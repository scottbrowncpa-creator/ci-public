# Recalculate xlsx files in real Excel (COM) on a temp copy; the original is replaced only when every formula has a cached value.
# Usage: powershell -File tools/excel_recalc.ps1 -Paths "<file1>" "<file2>"
#   -Paths "<f1>","<f2>" also works (under -File it arrives as one "f1,f2" string and is split here).
#   -PdfDir <folder> also exports each recalculated book to <folder>\<name>.pdf (all sheets, their print setup); no PDF = FAILED.
#   -CheckOnly validates the paths, python and the checker, prints "ok <path>"; Excel is not touched.
# Per file: copy to %TEMP%, Excel opens the copy, CalculateFull, forced Save (Excel skips saving a book it
# thinks is unchanged), close; then tools/xlsx_cached_check.py must find a cached value in every formula cell.
# Only then is the original overwritten, and only if its bytes did not change meanwhile; bytes compared after.
#   stdout "recalced <path>"   all of that passed and no formula returns an error.
#   stderr "ERRORS <path>: .." written back with every value cached, but some formulas return #REF!, #DIV/0! etc.
#   stderr "FAILED ..."        not recalced; the original is untouched unless the message says otherwise.
# Exit 0 = every file recalced clean; 2 = no failures but some ERRORS; 1 = anything FAILED.
# Excel: records the EXCEL PIDs, starts its own instance, finds its PID by window handle, and only ever quits or
# kills that one. An Excel already open is never touched. If Excel drops the connection (RPC_E_DISCONNECTED, or
# the process died: 0x800706BA / 0x800706BE) the file is retried once with a new Excel, noted on stderr.
# .xlsx/.xlsm only: the check reads them with openpyxl, so .xls/.xlsb are refused before Excel starts.
# Tests: python tools/test_excel_recalc.py [--live]
# Plain param block on purpose: extra args after -Paths land in $args and are merged below.
param(
  [string[]]$Paths,
  [switch]$CheckOnly
)
# -PdfDir is read from $args, not declared: a declared second string param would swallow the second positional path.
$PdfDir = $null
$rest = @()
for ($i = 0; $i -lt $args.Count; $i++) {
  if ($args[$i] -is [string] -and $args[$i] -ieq '-PdfDir' -and $i + 1 -lt $args.Count) { $PdfDir = [string]$args[$i + 1]; $i++ } else { $rest += ,$args[$i] }
}
$args = $rest
$ErrorActionPreference = 'Stop'

function Fail([string]$msg) { [Console]::Error.WriteLine($msg); exit 1 }
function Note([string]$msg) { [Console]::Error.WriteLine("note: $msg") }

# An arg that exists as-is is one file (commas allowed in names); otherwise split it on commas.
$files = @()
$bad = @()
foreach ($a in (@($Paths) + @($args))) {
  if ([string]::IsNullOrWhiteSpace($a)) { continue }
  if (Test-Path -LiteralPath $a -PathType Leaf) { $files += (Resolve-Path -LiteralPath $a).ProviderPath; continue }
  foreach ($piece in ($a -split ',')) {
    $p = $piece.Trim().Trim('"').Trim("'")
    if (-not $p) { continue }
    if (Test-Path -LiteralPath $p -PathType Leaf) { $files += (Resolve-Path -LiteralPath $p).ProviderPath }
    else { $bad += $p }
  }
}
if ($bad.Count) { Fail ("FAILED: file not found: " + ($bad -join ' | ') + " - nothing recalced") }
if (-not $files.Count) { Fail "FAILED: no paths given - usage: -Paths <file1> <file2>" }
$odd = @($files | Where-Object { @('.xlsx', '.xlsm') -notcontains [IO.Path]::GetExtension($_).ToLower() })
if ($odd.Count) { Fail ("FAILED: only .xlsx/.xlsm can be checked for cached values: " + ($odd -join ' | ') + " - nothing recalced") }

$checker = Join-Path $PSScriptRoot 'xlsx_cached_check.py'
if (-not (Test-Path -LiteralPath $checker -PathType Leaf)) { Fail "FAILED: $checker not found - no recalc could be verified, nothing recalced" }
$py = @(Get-Command python -CommandType Application -ErrorAction SilentlyContinue)
if (-not $py.Count) { Fail "FAILED: python not on PATH - xlsx_cached_check.py needs it, nothing recalced" }
$py = $py[0].Source

if ($PdfDir) {   # Excel needs a full path; relative would land in Excel's own current folder
  if (-not [IO.Path]::IsPathRooted($PdfDir)) { $PdfDir = Join-Path (Get-Location).ProviderPath $PdfDir }
  $PdfDir = [IO.Path]::GetFullPath($PdfDir)
  New-Item -ItemType Directory -Force -Path $PdfDir | Out-Null
}
if ($CheckOnly) { foreach ($f in $files) { Write-Output "ok $f" }; exit 0 }

# RPC_E_DISCONNECTED, RPC_S_SERVER_UNAVAILABLE, RPC_S_CALL_FAILED: the Excel being driven is gone or cut loose
$DROPPED = '80010108|800706BA|800706BE'
$script:xl = $null
$script:ours = @()
$script:keepTmp = $false
try {
  Add-Type -Namespace ExcelRecalc -Name Win32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);'
} catch { Note "window-handle lookup unavailable ($($_.Exception.Message)); the new Excel is matched by command line" }

function Test-Dropped($rec) {
  for ($e = $rec.Exception; $e; $e = $e.InnerException) {
    if (('{0:X8}' -f $e.HResult) -match "^($DROPPED)$" -or $e.Message -match "0x($DROPPED)") { return $true }
  }
  return $false
}

# An Excel that died does not always raise an RPC error: PowerShell reads its properties as $null
# ("You cannot call a method on a null-valued expression"). So ask the process and the object directly.
function Test-ExcelAlive {
  foreach ($o in $script:ours) { if (-not (Get-Process -Id $o.Id -ErrorAction SilentlyContinue)) { return $false } }
  try { return ($null -ne $script:xl.Workbooks) } catch { return $false }
}

# Lost = the Excel this run was driving is gone or cut loose (not a bad file, not a failed start).
function Test-Lost($rec) {
  if (Test-Dropped $rec) { return $true }
  return [bool]($script:xl -and -not (Test-ExcelAlive))
}

function Get-ExcelIds { @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) }

function Start-Excel {
  $before = Get-ExcelIds
  try { $app = New-Object -ComObject Excel.Application } catch { throw "cannot start Excel: $($_.Exception.Message)" }
  [uint32]$xpid = 0
  try { [void][ExcelRecalc.Win32]::GetWindowThreadProcessId([IntPtr][int64]$app.Hwnd, [ref]$xpid) } catch { $xpid = 0 }
  if ($xpid -and ($before -contains [int]$xpid)) {
    # COM handed back an Excel that was already running: never hide it, never quit it
    try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) } catch {}
    throw "Excel handed back an instance that was already open (pid $xpid) - not driving it"
  }
  if (-not $xpid) {
    $new = @(Get-ExcelIds | Where-Object { $before -notcontains $_ } | Where-Object {
        (Get-CimInstance Win32_Process -Filter "ProcessId=$_" -ErrorAction SilentlyContinue).CommandLine -match '/automation' })
    if ($new.Count -eq 1) { $xpid = [uint32]$new[0] }
  }
  if ($xpid) {
    $proc = Get-Process -Id ([int]$xpid) -ErrorAction SilentlyContinue
    $script:ours += [pscustomobject]@{ Id = [int]$xpid; Start = $(if ($proc) { $proc.StartTime } else { $null }) }
  } else {
    Note "could not tell which EXCEL process this run started; it gets Quit but is never force-killed"
  }
  $script:xl = $app
  $app.Visible = $false
  $app.DisplayAlerts = $false
  Start-Sleep -Seconds 8   # a new Excel rejects calls for a few seconds (RPC_E_CALL_REJECTED)
}

function Stop-Excel {
  if ($script:xl) {
    try { $script:xl.Quit() } catch {}
    try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:xl) } catch {}
    $script:xl = $null
  }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
  foreach ($o in $script:ours) {
    $proc = Get-Process -Id $o.Id -ErrorAction SilentlyContinue
    if (-not $proc -or $proc.ProcessName -ne 'EXCEL' -or ($o.Start -and $proc.StartTime -ne $o.Start)) { continue }  # gone, or the id was reused
    if ($proc.WaitForExit(15000)) { continue }
    try { Stop-Process -Id $o.Id -Force -ErrorAction Stop; Note "the Excel this run started (pid $($o.Id)) did not quit in 15s - killed it" }
    catch { Note "the Excel this run started (pid $($o.Id)) did not quit and could not be killed: $($_.Exception.Message)" }
  }
  $script:ours = @()
}

function Invoke-Recalc([string]$work, [string]$pdf) {
  $wb = $null
  try {
    $err = $null
    for ($i = 1; $i -le 3 -and -not $wb; $i++) {
      try { $wb = $script:xl.Workbooks.Open($work, 0, $false) }
      catch { if (Test-Lost $_) { throw }; $err = $_.Exception.Message; Start-Sleep -Seconds 3 }
    }
    if (-not $wb) { throw "Excel could not open it (3 tries): $err" }
    Start-Sleep -Seconds 5   # openpyxl books ask for a full calc on load; let that finish first
    $script:xl.CalculateFull()
    $done = $false
    for ($i = 0; $i -lt 120; $i++) {
      $state = $null
      try { $state = $script:xl.CalculationState } catch { if (Test-Lost $_) { throw } }
      if ($state -eq 0) { $done = $true; break }
      if ($null -eq $state -and -not (Test-ExcelAlive)) { throw "Excel stopped answering during the calculation" }
      Start-Sleep -Seconds 1
    }
    if (-not $done) { throw "calculation did not finish in 120s" }
    Start-Sleep -Seconds 3
    $wb.Saved = $false   # force the write: Excel skips Save on a book it thinks is unchanged
    $wb.Save()
    if ($pdf) {
      if (Test-Path -LiteralPath $pdf) { Remove-Item -LiteralPath $pdf -Force }
      $wb.ExportAsFixedFormat(0, $pdf)
      if (-not (Test-Path -LiteralPath $pdf)) { throw "Excel reported no error but wrote no PDF at $pdf" }
    }
    $wb.Close($false)
    $wb = $null
  } finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
  }
}

function Invoke-Check([string]$file) {
  $ErrorActionPreference = 'Continue'   # a stray stderr line from python must not throw here
  $out = @(& $py $checker $file 2>$null)
  $code = $LASTEXITCODE
  $ErrorActionPreference = 'Stop'
  $line = @($out | Where-Object { "$_".StartsWith('{') }) | Select-Object -Last 1
  if (-not $line) { throw "xlsx_cached_check.py gave no answer (exit $code) - original not overwritten" }
  $r = $line | ConvertFrom-Json
  $r | Add-Member -NotePropertyName code -NotePropertyValue $code -Force
  return $r
}

function Recalc-File([string]$path, [string]$dir) {
  $name = [IO.Path]::GetFileName($path)
  $keep = Join-Path $dir 'before'
  New-Item -ItemType Directory -Force -Path $keep | Out-Null
  $orig = Join-Path $keep $name
  Copy-Item -LiteralPath $path -Destination $orig -Force
  $origHash = (Get-FileHash -LiteralPath $orig -Algorithm SHA256).Hash
  $work = $null
  for ($attempt = 1; $attempt -le 2; $attempt++) {
    $adir = Join-Path $dir "try$attempt"
    New-Item -ItemType Directory -Force -Path $adir | Out-Null
    $work = Join-Path $adir $name
    Copy-Item -LiteralPath $orig -Destination $work -Force
    try {
      if (-not $script:xl) { Start-Excel }
      Invoke-Recalc $work $(if ($PdfDir) { Join-Path $PdfDir ([IO.Path]::ChangeExtension($name, '.pdf')) } else { $null })
      break
    } catch {
      if (-not (Test-Lost $_)) { throw }
      $why = if (Test-Dropped $_) { $_.Exception.Message.Trim() } else { 'its process died or stopped answering' }
      Stop-Excel
      if ($attempt -eq 2) { throw "Excel dropped the connection again after 1 retry: $why" }
      Note "Excel dropped the connection on $name ($why); retried once with a new Excel"
    }
  }

  $r = Invoke-Check $work
  if ($r.error) { throw "the saved copy could not be read back to verify it: $($r.error) - original not overwritten" }
  if ($null -eq $r.formulas -or $null -eq $r.uncached -or $null -eq $r.errors) { throw "xlsx_cached_check.py answer not understood - original not overwritten" }
  if ($r.uncached -gt 0) {
    $more = if ($r.uncached_more) { " and $($r.uncached_more) more" } else { '' }
    throw ("{0} of {1} formula cells have no cached value after the save: {2}{3} - original not overwritten" -f $r.uncached, $r.formulas, ($r.uncached_cells -join ', '), $more)
  }
  if (@(0, 2) -notcontains $r.code) { throw "xlsx_cached_check.py exited $($r.code) - original not overwritten" }

  # bytes, not timestamps: a FileInfo reads its mtime lazily, so a snapshot taken with Get-Item saw the rewrite too
  if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $origHash) {
    throw "the file changed on disk while Excel had the copy - not overwritten; rerun"
  }
  $hash = (Get-FileHash -LiteralPath $work -Algorithm SHA256).Hash
  $cerr = $null
  for ($i = 1; $i -le 3; $i++) {
    try { Copy-Item -LiteralPath $work -Destination $path -Force; $cerr = $null; break }
    catch { $cerr = $_.Exception.Message; Start-Sleep -Seconds 2 }
  }
  $destHash = ''
  try { $destHash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash } catch {}
  if ($destHash -ne $hash) {
    if ($destHash -eq $origHash) { throw "could not write the recalculated copy back (3 tries: $cerr) - original untouched" }
    $script:keepTmp = $true
    throw "write-back went wrong and the file may be damaged - pre-run copy kept at $orig, recalculated copy at $work"
  }
  return $r
}

$tmpRoot = Join-Path $env:TEMP ("excel_recalc_" + [guid]::NewGuid().ToString('N'))
$failed = 0
$errored = 0
try {
  $n = 0
  foreach ($f in $files) {
    $n++
    try {
      $r = Recalc-File $f (Join-Path $tmpRoot ([string]$n))
      if ($r.errors -gt 0) {
        $errored++
        $more = if ($r.errors_more) { " and $($r.errors_more) more" } else { '' }
        [Console]::Error.WriteLine("ERRORS ${f}: saved with a value in every formula, but $($r.errors) return an error: " + ($r.error_cells -join ', ') + $more)
      } else {
        Write-Output "recalced $f"
      }
    } catch {
      $failed++
      [Console]::Error.WriteLine("FAILED ${f}: $($_.Exception.Message)")
    }
  }
} finally {
  Stop-Excel
  if ($script:keepTmp) { Note "kept $tmpRoot for recovery" }
  else {
    Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $tmpRoot) { Note "could not remove $tmpRoot" }
  }
}
if ($failed) { exit 1 }
if ($errored) { exit 2 }
exit 0
