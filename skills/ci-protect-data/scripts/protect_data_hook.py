#!/usr/bin/env python3
"""protect_data_hook.py - Claude Code hook that enforces ci-protect-data on every file read, not only when the skill loads.

  python protect_data_hook.py pre             PreToolUse (Read, Grep, Bash, PowerShell, PDF tools), payload on stdin
  python protect_data_hook.py prompt          UserPromptSubmit: warns when a pasted message holds bank, student or SSN numbers
  python protect_data_hook.py --scan <file>   header/keyword scan, used for .pdf and .xls and when the full check runs long

On a read of .xlsx .xlsm .xls .csv .pdf:
  student names/IDs/MARSS/state IDs/birth dates, or employee bank account/routing numbers  -> DENY: use "<file> MASKED.xlsx"
  SSNs only                                                                                -> warn, never block
  a "* MASKED.*" file                                                                      -> pass
xlsx/xlsm/csv go through mask_sensitive.py --check (same folder). It finds columns by their header names, so a roster with
plain First/Last Name headers and nothing that says Student gets through. PDFs get a text scan of the first 40 pages for
strong patterns only (routing word plus 3+ valid routing numbers; a student ID word plus 3+ long ID numbers; Student Name
plus a birth date header). .xls headers are read with xlrd.

Fail closed: a payload it cannot read, a check that errors, times out or cannot open the file is DENY. One stated carve-out,
reported as a warning instead: a PDF it cannot read the text of (no pypdf installed, scanned with no text layer, broken
file), because PDF reads are mostly statements and packets and denying them all would get the hook switched off.
Caps (40 PDF pages, the full-check time limit, 8 files per call) are reported when hit.
Shell commands: checks every existing file a command names unless the command only lists, copies, moves, opens it for the
person (Start-Process, explorer) or runs mask_sensitive.py. Paths with an unknown variable are reported, not checked.
Off switch for a person, not Claude: env CI_PROTECT_DATA=off (no checks) or =warn (warn instead of deny).
The plugin copy steps aside when the project registers its own copy (.claude/hooks/protect_data_guard.py), so it runs once.
Limit: it stops accidental reads. A script that builds a file name at run time, or a roster with plain headers, gets past it.
"""
import json
import os
import re
import sys

EXT_RX = re.compile(r"\.(?:xlsx|xlsm|xls|csv|pdf)(?![\w])", re.I)
HERE = os.path.dirname(os.path.abspath(__file__))
MASK = os.path.join(HERE, "mask_sensitive.py")
MODE = (os.environ.get("CI_PROTECT_DATA") or "on").strip().lower()
try:                                                               # full mask_sensitive check, per file
    CHECK_SECS = min(30.0, max(1.0, float(os.environ.get("CI_PROTECT_DATA_SECS") or 20)))
except ValueError:                                                 # a bad value must not crash the guard open
    CHECK_SECS = 20.0
SCAN_SECS = 15                                                     # header/keyword scan, per file
DEADLINE = 45                                                      # whole call; settings timeout must be above this
MAX_FILES = 8
PDF_PAGES = 40
KIND_TEXT = {"bank": "employee bank account/routing numbers", "student": "student names, IDs or birth dates",
             "ssn": "SSN-like numbers"}
NOWIN = 0x08000000 if os.name == "nt" else 0                       # CREATE_NO_WINDOW

# Commands that never put file contents in front of Claude.
EXEMPT = {"ls", "dir", "gci", "get-childitem", "test-path", "get-item", "stat", "file", "cp", "copy", "copy-item",
          "cpi", "mv", "move", "move-item", "mi", "ren", "rename", "rename-item", "rni", "rm", "del", "erase",
          "remove-item", "ri", "mkdir", "md", "new-item", "start", "start-process", "saps", "ii", "invoke-item",
          "explorer", "echo", "write-host", "write-output", "attrib", "robocopy", "xcopy", "icacls",
          "get-filehash", "certutil", "touch", "where", "which", "resolve-path", "split-path", "join-path"}
CD = {"cd", "chdir", "set-location", "sl", "pushd", "push-location"}
PY = {"python", "python3", "py", "pythonw"}
SEP = re.compile(r"&&|\|\||[;|\n]")
BOUND = set(" \t\"'=(,")


def is_masked(p):
    return bool(re.search(r"\sMASKED$", os.path.splitext(os.path.basename(p))[0], re.I))


# ------------------------------------------------------------------ paths
def resolve(p, cwd, var):
    """Absolute path, or None when it still holds a variable this hook cannot fill in."""
    p = p.strip().strip("\"'").replace("\\ ", " ")
    p = re.sub(r"\$env:(\w+)", lambda m: os.environ.get(m.group(1), m.group(0)), p, flags=re.I)
    p = re.sub(r"\$\{?(\w+)\}?", lambda m: var.get(m.group(1), os.environ.get(m.group(1), m.group(0))), p)
    p = re.sub(r"%(\w+)%", lambda m: os.environ.get(m.group(1), m.group(0)), p)
    if re.search(r"\$\w|%\w+%", p):
        return None
    m = re.match(r"^/([a-zA-Z])/(.*)$", p)                 # Git Bash /c/... form
    if m:
        p = m.group(1).upper() + ":/" + m.group(2)
    p = os.path.expanduser(p)
    if not os.path.isabs(p):
        p = os.path.join(cwd, p)
    return os.path.normpath(p)


def candidates(text, cwd, var, notes):
    """Existing target files named in text. For each extension hit, the longest left-extension that is a real file wins,
    so quoted paths, bare tokens, -c code strings and unquoted paths with spaces all resolve."""
    found = []
    text = text.replace('\\"', '"').replace("\\'", "'")
    for m in EXT_RX.finditer(text):
        j = m.end()
        line_start = text.rfind("\n", 0, m.start()) + 1
        starts = [line_start] + [i + 1 for i in range(max(line_start, j - 320), m.start()) if text[i] in BOUND]
        starts = sorted(set(starts))
        nearest = text[max(starts):j]
        best = None
        for s in starts:                                    # longest first
            c = text[s:j].strip().strip("\"'(=,").strip()
            r = resolve(c, cwd, var) if c else None
            if r and os.path.isfile(r):
                best = r
                break
        if best:
            found.append(best)
            continue
        tok = nearest.strip().strip("\"'")
        if any(ch in tok for ch in "*?"):
            import glob
            r = resolve(tok, cwd, var)
            hits = sorted(glob.glob(r)) if r else []
            if len(hits) > MAX_FILES:
                notes.append(f"{tok}: {len(hits)} files match")
            found.extend(h for h in hits if os.path.isfile(h))
        elif resolve(tok, cwd, var) is None:
            notes.append(f"could not resolve {tok}, not checked")
    return found


def shell_paths(cmd, cwd, notes):
    if not EXT_RX.search(cmd):
        return []
    out, var = [], {}
    for seg in SEP.split(cmd):
        s = seg.strip()
        a = re.match(r"^\$?(\w+)\s*=\s*(\"[^\"]*\"|'[^']*'|\S+)\s*$", s)   # NAME=value / $Name = "value"
        if a:
            var[a.group(1)] = a.group(2).strip("\"'")
            continue
        toks = s.split()
        while toks and (toks[0] in ("&", ".", "sudo", "time", "env") or re.match(r"^\w+=\S*$", toks[0])):
            toks.pop(0)
        if not toks:
            continue
        verb = os.path.basename(toks[0].strip("\"'").replace("\\", "/")).lower()
        if verb.endswith(".exe"):
            verb = verb[:-4]
        if verb in CD and len(toks) > 1:
            t = resolve(" ".join(toks[1:]), cwd, var)
            if t and os.path.isdir(t):
                cwd = t
            continue
        if verb in EXEMPT:
            continue
        if re.search(r"mask_sensitive\.py|protect_data_hook\.py|protect_data_guard\.py", s, re.I) and (
                verb in PY or verb.endswith(".py")):
            continue
        if var:                                             # "$F" set earlier in the same command
            s = re.sub(r"\$\{?(\w+)\}?", lambda m: var.get(m.group(1), m.group(0)), s)
        out += candidates(s, cwd, var, notes)
    return out


def field_paths(name, inp, cwd, notes):
    out = []
    vals = []
    for v in inp.values():
        if isinstance(v, str):
            vals.append(v)
        elif isinstance(v, list):
            vals += [x for x in v if isinstance(x, str)]
    for v in vals:
        if EXT_RX.search(v):
            r = resolve(v, cwd, {})
            if r and os.path.isfile(r):
                out.append(r)
    if name == "Grep":
        g = str(inp.get("glob") or "") + " " + str(inp.get("type") or "")
        if re.search(r"csv", g, re.I) and not (inp.get("path") and os.path.isfile(str(inp.get("path")))):
            notes.append("Grep across csv files is not checked by protect-data; matching rows can carry raw values")
    return out


# ------------------------------------------------------------------ scanning
SSN_TEXT = re.compile(r"(?<![\d-])(\d{3})([- ])(\d{2})\2(\d{4})(?![\d-])")
NINE = re.compile(r"(?<![\d-])\d{9}(?![\d-])")
BANK_WORD = re.compile(r"routing|\baba\b|\brtn\b|transit\s*(?:no|num|#)", re.I)
STU_ID_WORD = re.compile(r"\bmarss\b|state\s*student\s*id\b|student\s*(?:id|#|number|no\b)|\beduid\b|\bssid\b", re.I)
STU_NAME = re.compile(r"student\s*name|student\s*(?:first|last)", re.I)
DOB_WORD = re.compile(r"\bdob\b|date\s*of\s*birth|birth\s*date", re.I)
LONGNUM = re.compile(r"(?<![\d,.$-])\d{6,13}(?![\d,.])")


def aba_ok(d):
    if len(d) != 9 or len(set(d)) == 1:
        return False
    n = [int(c) for c in d]
    return (3 * (n[0] + n[3] + n[6]) + 7 * (n[1] + n[4] + n[7]) + (n[2] + n[5] + n[8])) % 10 == 0


def text_kinds(text):
    """Kinds found in free text (PDF pages). Strong patterns only, so statements and board packets pass."""
    k = set()
    if BANK_WORD.search(text) and sum(1 for d in NINE.findall(text) if aba_ok(d)) >= 3:
        k.add("bank")
    if (STU_ID_WORD.search(text) and len(LONGNUM.findall(text)) >= 3) or (STU_NAME.search(text) and DOB_WORD.search(text)):
        k.add("student")
    if SSN_TEXT.search(text):
        k.add("ssn")
    return k


def header_kinds(rows_by_sheet):
    sys.path.insert(0, HERE)
    import mask_sensitive as ms
    k = set()
    for rows in rows_by_sheet:
        hi = ms.header_row(rows)
        if hi is not None:
            k |= {c for c in (ms.classify(h) for h in rows[hi]) if c}
    return k


def scan(path):
    """dict(kinds=[...], notes=[...], unchecked=reason-or-None). Raises on anything it cannot read (caller denies)."""
    ext = os.path.splitext(path)[1].lower()
    notes = []
    if ext == ".pdf":
        try:
            import pypdf
        except ImportError:
            return {"kinds": [], "notes": notes, "unchecked": "no PDF reader (pypdf) installed"}
        try:
            r = pypdf.PdfReader(path)
            if r.is_encrypted:
                r.decrypt("")
            n = len(r.pages)
            text = "\n".join((r.pages[i].extract_text() or "") for i in range(min(n, PDF_PAGES)))
        except Exception as e:
            return {"kinds": [], "notes": notes, "unchecked": f"PDF text could not be read ({type(e).__name__})"}
        if n > PDF_PAGES:
            notes.append(f"first {PDF_PAGES} of {n} pages checked")
        if len(text.strip()) < 20:
            return {"kinds": [], "notes": notes, "unchecked": "no text layer (scanned?)"}
        return {"kinds": sorted(text_kinds(text)), "notes": notes, "unchecked": None}
    if ext == ".xls":
        import xlrd                                      # ImportError -> caller denies (fail closed)
        wb = xlrd.open_workbook(path, on_demand=True)
        sheets, ssn = [], False
        for sh in wb.sheets():
            sheets.append([sh.row_values(i) for i in range(min(10, sh.nrows))])
            for i in range(min(sh.nrows, 20000)):
                if any(isinstance(v, str) and SSN_TEXT.search(v) for v in sh.row_values(i)):
                    ssn = True
                    break
        k = header_kinds(sheets) | ({"ssn"} if ssn else set())
        return {"kinds": sorted(k), "notes": ["headers checked (.xls)"], "unchecked": None}
    if ext == ".csv":
        import csv
        with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
            rows = [r for _, r in zip(range(10), csv.reader(f))]
        return {"kinds": sorted(header_kinds([rows])), "notes": ["headers only"], "unchecked": None}
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheets = [[list(r) for r in ws.iter_rows(max_row=10, values_only=True)] for ws in wb.worksheets]
    wb.close()
    return {"kinds": sorted(header_kinds(sheets)), "notes": ["headers only"], "unchecked": None}


def run(argv, secs):
    import subprocess
    return subprocess.run([sys.executable] + argv, capture_output=True, text=True, encoding="utf-8", errors="replace",
                          timeout=secs, stdin=subprocess.DEVNULL, creationflags=NOWIN)


def run_scan(path):
    import subprocess
    try:
        r = run([os.path.abspath(__file__), "--scan", path], SCAN_SECS)
    except subprocess.TimeoutExpired:
        return {"error": f"scan ran past {SCAN_SECS}s"}
    try:
        res = json.loads(r.stdout.strip().splitlines()[-1])
    except Exception:
        return {"error": "scan failed: " + (r.stderr or r.stdout or "no output").strip()[-160:]}
    return res


MASK_LINE = re.compile(r"^(?:WOULD MASK|MASKED):\s*(.+)$", re.M)


def check_file(path):
    """dict(kinds, counts, notes, unchecked, error) for one file."""
    if is_masked(path):
        return {"kinds": []}
    ext = os.path.splitext(path)[1].lower()
    if ext not in (".xlsx", ".xlsm", ".csv"):
        return run_scan(path)
    if not os.path.isfile(MASK):
        return {"error": "mask_sensitive.py is missing next to the hook"}
    import subprocess
    try:
        r = run([MASK, path, "--check"], CHECK_SECS)
    except subprocess.TimeoutExpired:
        res = run_scan(path)
        res.setdefault("notes", []).append(f"full check passed the {int(CHECK_SECS)}s cap: headers only, SSN text not scanned")
        return res
    out = r.stdout or ""
    if r.returncode == 0:
        return {"kinds": ["ssn"] if re.search(r"^WARNING:.*SSN", out, re.M) else []}
    if r.returncode == 1:
        m = MASK_LINE.search(out)
        if not m:
            return {"error": "mask check output not understood"}
        counts = {}
        for n, kind in re.findall(r"(\d+)\s+([\w-]+)", m.group(1)):
            counts[kind] = counts.get(kind, 0) + int(n)
        if not counts:
            return {"error": "mask check output not understood"}
        kinds = set()
        for kind in counts:
            kinds.add("ssn" if kind.startswith("ssn") else kind if kind in ("bank", "student") else "unknown:" + kind)
        if re.search(r"^WARNING:.*SSN", out, re.M):
            kinds.add("ssn")
        return {"kinds": sorted(kinds), "counts": counts}
    return {"error": (out.strip().splitlines() or ["mask check exit %d" % r.returncode])[0][:160]}


# ------------------------------------------------------------------ output
def emit(obj):
    print(json.dumps(obj))
    sys.exit(0)


def deny(reason):
    emit({"hookSpecificOutput": {"hookEventName": "PreToolUse", "permissionDecision": "deny",
                                 "permissionDecisionReason": "protect-data: " + reason}})


def warn(text, event="PreToolUse"):
    emit({"systemMessage": "protect-data: " + text.split("\n")[0][:300],
          "hookSpecificOutput": {"hookEventName": event, "additionalContext": "protect-data: " + text}})


def stepped_aside():
    """True when this is the plugin copy and the project registers its own copy of the guard."""
    proj = os.environ.get("CLAUDE_PROJECT_DIR")
    if not proj:
        return False
    here, pj = os.path.normcase(os.path.abspath(__file__)), os.path.normcase(os.path.abspath(proj))
    if here.startswith(pj.rstrip("\\/") + os.sep):
        return False
    try:
        reg = "protect_data_guard" in open(os.path.join(proj, ".claude", "settings.json"), encoding="utf-8-sig").read()
    except OSError:
        reg = False
    return reg and os.path.isfile(os.path.join(proj, ".claude", "hooks", "protect_data_guard.py"))


def pre(raw):
    import time
    t0 = time.time()
    try:
        data = json.loads(raw)
        name = str(data.get("tool_name") or "")
        inp = data.get("tool_input")
        if not isinstance(inp, dict):
            raise ValueError("tool_input is not an object")
        cwd = str(data.get("cwd") or os.getcwd())
    except Exception as e:
        deny(f"could not read the tool payload ({type(e).__name__}); not cleared.")
    notes = []
    if name in ("Bash", "PowerShell"):
        paths = shell_paths(str(inp.get("command") or ""), cwd, notes)
    else:
        paths = field_paths(name, inp, cwd, notes)
    seen, files = set(), []
    for p in paths:
        k = os.path.normcase(p)
        if k not in seen:
            seen.add(k)
            files.append(p)
    if len(files) > MAX_FILES:
        deny(f"this call names {len(files)} files with possible sensitive data; the check covers {MAX_FILES} per call. "
             "Split it, or mask the files first.")
    results = {}
    if files:
        import concurrent.futures as cf
        ex = cf.ThreadPoolExecutor(max_workers=4)
        futs = {ex.submit(check_file, p): p for p in files}
        done, pending = cf.wait(futs, timeout=max(1, DEADLINE - (time.time() - t0)))
        for f in done:
            try:
                results[futs[f]] = f.result()
            except Exception as e:
                results[futs[f]] = {"error": f"check crashed ({type(e).__name__})"}
        for f in pending:
            results[futs[f]] = {"error": f"check did not finish inside {DEADLINE}s"}
        try:
            ex.shutdown(wait=False, cancel_futures=True)
        except TypeError:                                          # Python 3.8 has no cancel_futures
            ex.shutdown(wait=False)
    blocked, warned, quiet = [], [], list(notes)
    for p in files:
        r = results[p]
        base = os.path.basename(p)
        extra = ("; " + "; ".join(r["notes"])) if r.get("notes") else ""
        if r.get("error"):
            blocked.append(f"{base}: could not be checked ({r['error']}), so it is not cleared{extra}")
            continue
        if r.get("unchecked"):
            warned.append(f"{base}: not checked, {r['unchecked']}. Look before relying on it{extra}")
            continue
        kinds = r.get("kinds") or []
        bad = [k for k in kinds if k != "ssn"]
        if bad:
            what = ", ".join(KIND_TEXT.get(k, k) for k in bad)
            stem, ext = os.path.splitext(p)
            if ext.lower() in (".xlsx", ".xlsm", ".csv"):
                fix = (f'Make a masked copy: python "{MASK}" "{p}" then read "{os.path.basename(stem)} MASKED{ext}" '
                       "(the original is not changed).")
            else:
                fix = ("mask_sensitive cannot mask a " + ext.lstrip(".").upper() + ". Use the Excel export and mask that, "
                       "or have the person blank those columns.")
            blocked.append(f"{base} has {what}{extra}. {fix}")
        elif "ssn" in kinds:
            warned.append(f"{base} has SSN-like numbers. Read allowed (SSNs warn only); never repeat them, last 4 at most{extra}")
        elif r.get("notes"):
            quiet.append(f"{base}: clear, but {'; '.join(r['notes'])}")
    if blocked and MODE != "warn":
        deny("not read. " + " | ".join(blocked + warned + quiet) + " | A person can set CI_PROTECT_DATA=warn to allow it.")
    if blocked or warned:
        warn(" | ".join(blocked + warned + quiet))
    if quiet:                                              # caps and limits: Claude is told, the screen is not
        emit({"hookSpecificOutput": {"hookEventName": "PreToolUse", "additionalContext": "protect-data: " + " | ".join(quiet)}})
    sys.exit(0)


# ------------------------------------------------------------------ prompt (warn only)
ACCT = re.compile(r"(?:account|acct)\s*(?:number|no\.?|#)\s*[:#]?\s*\d{6,17}\b", re.I)
STU_PASTE = re.compile(r"(?:student\s*(?:id|#|number|no\.?)|\bmarss(?:\s*(?:id|#|number|no\.?))?|\bssid|\beduid|"
                       r"state\s*(?:student\s*)?id)\s*[:#]?\s*\d{5,13}\b", re.I)


def prompt_kinds(text):
    k = set()
    abas = [m for m in NINE.finditer(text) if aba_ok(m.group(0))]
    if abas and BANK_WORD.search(text):
        near = any(BANK_WORD.search(text[max(0, m.start() - 80):m.end() + 80]) for m in abas)
        if near or len(abas) >= 2:
            k.add("bank")
    if ACCT.search(text):
        k.add("bank")
    if STU_PASTE.search(text) or (STU_ID_WORD.search(text) and len(LONGNUM.findall(text)) >= 3):
        k.add("student")
    if SSN_TEXT.search(text):
        k.add("ssn")
    return k


def prompt(raw):
    try:
        data = json.loads(raw)
        text = data.get("prompt")
        if not isinstance(text, str):
            raise ValueError("prompt is not text")
        kinds = prompt_kinds(text)
    except Exception as e:
        warn(f"could not check this message ({type(e).__name__}). Treat pasted numbers as possibly sensitive.",
             "UserPromptSubmit")
    if not kinds:
        sys.exit(0)
    what = ", ".join(KIND_TEXT[k] for k in sorted(kinds))
    warn(f"this message looks like it holds {what}. Do not repeat those numbers or names back: last 4 or Student 001 tags. "
         "At the end of the chat, remind once about /ci-wrapup.", "UserPromptSubmit")


def main():
    if len(sys.argv) >= 3 and sys.argv[1] == "--scan":
        try:
            print(json.dumps(scan(sys.argv[2])))
        except Exception as e:
            print(json.dumps({"error": f"{type(e).__name__}: {e}"[:200]}))
        return 0
    mode = sys.argv[1] if len(sys.argv) > 1 else "pre"
    try:
        raw = sys.stdin.buffer.read().decode("utf-8-sig", errors="replace").lstrip(chr(0xFEFF))
    except Exception as e:
        raw = None
        err = type(e).__name__
    if MODE == "off" or stepped_aside():
        return 0
    if raw is None:
        if mode == "prompt":
            warn(f"could not read this message ({err}).", "UserPromptSubmit")
        deny(f"could not read the tool payload ({err}); not cleared.")
    try:
        if mode == "prompt":
            prompt(raw)
        pre(raw)
    except SystemExit:
        raise
    except Exception as e:                               # fail closed
        if mode == "prompt":
            warn(f"check errored ({type(e).__name__}). Treat pasted numbers as possibly sensitive.", "UserPromptSubmit")
        deny(f"the check errored ({type(e).__name__}: {str(e)[:120]}); not cleared.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
