"""Ship check: math and accounting ties for an xlsx, and number ties for an html page built from it. Read-only.

    python tools/ship_check.py xlsx "<file.xlsx>" [--voice]
    python tools/ship_check.py html "<page.html>" --xlsx "<file.xlsx>" [--page-text "<rendered text file>"]
    python tools/ship_check.py docx "<file.docx>"      fingerprints (library author stamp, no Word app tag, build talk,
                                                       git hashes) plus the AI-tell list per paragraph; MED, stamps when clean

xlsx: reads formulas and cached values together (recalc first: tools/excel_recalc.ps1). Settles
  - every formula has a cached value and none is an error
  - every labeled Total / Subtotal row foots to the cells it sums (typed totals are compared to the cells above)
  - rows labeled Starting/Opening and Ending roll forward column to column, and Ending = Starting + Total in + Total out
  - rows labeled Check / Difference / Out of balance are zero
  - Debit and Credit columns are equal, overall and per JE / batch column when there is one
  - numbers stored as text, exact duplicate detail rows
  then runs tools/wb_lint.py for the house conventions and prints its findings after ours.
html: static page checks (title, viewport, placeholders, dark-mode block) plus, with --page-text (the page text as
  rendered in the browser), every $ figure on the page is looked up in the workbook. A figure the workbook does not
  carry is a finding: the reader sees two numbers for one thing.
The AI-tell find-list runs over authored cells by default (MED, reported, not blocking); --no-voice skips it.
Every mode also reports placeholder library metadata in docProps (creator "openpyxl" / "python-docx" / "Un-named", no Office
Application tag); `python tools/doc_fingerprint.py fix <file> --author "<name>"` clears them before the file goes out.

An xlsx is stamped only if it also passes the xlsx-ship-ready gate (tools/xlsx_gate.py: recalc, wb_lint, footing).
--force-reason "<why>" overrides that, reviewer only, logged to logs/gate.jsonl.
A clean pass (exit 0) is stamped in review/ship-check-ledger.jsonl with the file's SHA-256; --no-stamp skips that.
    python tools/ship_check.py stamped "<file>" [--xlsx "<xlsx>"]   exit 0 = these exact bytes passed before, 1 = not stamped
Ask `stamped` before anything goes to a client or auditor. Changed by one cell = new hash = run the check again.

Exit 0 = no HIGH, 1 = HIGH found, 3 = something could not be evaluated and nothing HIGH (never a clean result).
Caps report themselves on a TRUNCATED line.
"""

import argparse
import datetime as dt
import hashlib
import json
import os
import re
import subprocess
import sys
from pathlib import Path

import openpyxl

_WS = Path(__file__).resolve().parent.parent
# Outside the workspace (a staff plugin install) the folder next to scripts/ is wiped on every plugin update,
# so the ledger lives in the person's home folder instead.
_STATE = (_WS / "review") if (_WS / "CLAUDE.md").is_file() else (Path.home() / ".ci-tools")
LEDGER = Path(os.environ.get("SHIP_CHECK_LEDGER") or _STATE / "ship-check-ledger.jsonl")
# In the workspace a client-bound file also needs a data:validate-data pass on record for its bytes.
# Staff installs have no Data plugin, so off there. SHIP_CHECK_VALIDATE=0/1 forces it (tests).
VALIDATE_REQUIRED = os.environ.get("SHIP_CHECK_VALIDATE", "1" if _STATE == _WS / "review" else "0") == "1"
GATED = (".xlsx", ".xlsm", ".xls", ".html", ".htm")   # attachments that must carry a stamp to reach an outside recipient
TOL = 0.5          # dollars; a footing difference under 50 cents is rounding in display, not an error
MAX_LIST = 25
SEV = {"HIGH": 0, "MED": 1, "INFO": 2, "UNKNOWN": 3}
ERRORS = ("#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NULL!", "#NUM!")
SUM_RX = re.compile(r"^=\s*(?:SUM|SUBTOTAL\(\s*9\s*,)\s*\(?\s*\$?([A-Z]{1,3})\$?(\d+)\s*:\s*\$?([A-Z]{1,3})\$?(\d+)\s*\)+$", re.I)
TOTAL_LABEL = re.compile(r"^\s*(sub-?total|total|grand total)\b", re.I)
BLOCK_START = re.compile(r"^\s*(starting|opening|beginning)\b", re.I)   # a balance row that opens a block; never a component of its total
CHECK_LABEL = re.compile(r"^\s*(check\b|out of balance|s/b 0)", re.I)
CHECKS_HDR = re.compile(r"^\s*checks\b.*(\b(must|should) be (0|zero)\b|\bs/b 0\b|\ball 0\b)", re.I)   # block header
FLAG_HDR_RX = re.compile(r"^\s*client flags\b", re.I)          # client findings below a CHECKS block, not file errors
DIFF_BLOCK_RX = re.compile(r"^\s*differences\b.*\bexplained\b", re.I)   # reconciling block under CHECKS: nonzero is the explanation
DIFF_LABEL = re.compile(r"^\s*(difference|diff\b|variance to)", re.I)   # a difference line is allowed to be nonzero
SIGN_NOTE = re.compile(r"\b(negative|positive)\s*=|\(over\)|\bover the limit\b", re.I)   # label says what the sign means


class Report:
    def __init__(self):
        self.items = []
        self.counts = {}
        self.ran = {}

    def did(self, check, n=1):
        self.ran[check] = self.ran.get(check, 0) + n

    def add(self, sev, check, text):
        n = self.counts.get(check, 0)
        self.counts[check] = n + 1
        if n < MAX_LIST:
            self.items.append((sev, check, text))

    def render(self):
        out = []
        for check, n in self.counts.items():
            if n > MAX_LIST:
                out.append(f"TRUNCATED  {check}: showing {MAX_LIST} of {n}")
        by = sorted(self.items, key=lambda i: SEV[i[0]])
        return by, out

    def worst(self):
        sevs = {i[0] for i in self.items}
        if "HIGH" in sevs:
            return 1
        if "UNKNOWN" in sevs:
            return 3
        return 0


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def col_idx(letters):
    return openpyxl.utils.column_index_from_string(letters.upper())


def label_of(ws, row):
    """The row's label: the first text cell in the first three columns."""
    for c in range(1, 4):
        v = ws.cell(row=row, column=c).value
        if isinstance(v, str) and v.strip():
            return v.strip()
    return ""


def load_both(path):
    try:
        wf = openpyxl.load_workbook(path)
        wv = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:  # unreadable file is not a pass
        print(f"REFUSED: cannot read {path}: {e}")
        sys.exit(2)
    return wf, wv


def check_cached(wf, wv, rep):
    for ws in wf.worksheets:
        vs = wv[ws.title]
        missing = []
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    rep.did("formulas cached")
                    v = vs[c.coordinate].value
                    if v is None and re.search(r'IFERROR|IF\(|""', c.value, re.I):
                        rep.did("formulas that may return blank")   # cached "" reads back as empty; not an unknown
                    elif v is None:
                        missing.append(c.coordinate)
                    elif isinstance(v, str) and v in ERRORS:
                        rep.add("HIGH", "errors", f"{ws.title}!{c.coordinate} returns {v}")
        if missing:
            rep.add("UNKNOWN", "cached", f"{ws.title}: {len(missing)} formula cell(s) have no cached value ({', '.join(missing[:4])}...). "
                    "Recalc first: tools/excel_recalc.ps1 -Paths <file>. Ties below were skipped for these cells.")


def check_footing(wf, wv, rep):
    for ws in wf.worksheets:
        vs = wv[ws.title]
        for r in range(1, ws.max_row + 1):
            lab = label_of(ws, r)
            if not TOTAL_LABEL.match(lab):
                continue
            for c in range(2, ws.max_column + 1):
                f, v = ws.cell(row=r, column=c).value, vs.cell(row=r, column=c).value
                if isinstance(f, str) and f.startswith("="):
                    m = SUM_RX.match(f)
                    if not m or not num(v):
                        continue
                    c1, r1, c2, r2 = col_idx(m[1]), int(m[2]), col_idx(m[3]), int(m[4])
                    cells = [vs.cell(row=rr, column=cc).value for rr in range(r1, r2 + 1) for cc in range(c1, c2 + 1)]
                    if any(isinstance(x, str) and x.startswith("=") for x in cells):
                        continue
                    have = sum(x for x in cells if num(x))
                    rep.did("total cells footed")
                    if abs(have - v) > TOL:
                        rep.add("HIGH", "footing", f"{ws.title}!{ws.cell(row=r, column=c).coordinate} '{lab}' shows {v:,.2f}, "
                                f"its range {m[1]}{m[2]}:{m[3]}{m[4]} adds to {have:,.2f}")
                elif num(f):
                    above = []
                    rr = r - 1
                    while rr >= 1 and num(ws.cell(row=rr, column=c).value):
                        # a Beginning/Opening/Starting row or an earlier Total is where this block starts,
                        # not a component of it
                        above_lab = label_of(ws, rr)
                        if BLOCK_START.match(above_lab) or TOTAL_LABEL.match(above_lab):
                            break
                        above.append(ws.cell(row=rr, column=c).value)
                        rr -= 1
                    if above:
                        rep.did("total cells footed")
                        have = sum(above)
                        if abs(have - f) > TOL:
                            rep.add("HIGH", "footing", f"{ws.title}!{ws.cell(row=r, column=c).coordinate} '{lab}' typed {f:,.2f}, "
                                    f"the {len(above)} cells above add to {have:,.2f}")
                        else:
                            rep.add("INFO", "footing", f"{ws.title}!{ws.cell(row=r, column=c).coordinate} '{lab}' is typed but ties; make it a formula")


def _rows(ws, starts):
    return [r for r in range(1, ws.max_row + 1) if label_of(ws, r).lower().startswith(starts)]


def check_rollforward(wv, rep):
    for vs in wv.worksheets:
        opens, ends = _rows(vs, ("starting", "opening", "beginning")), _rows(vs, ("ending", "closing"))
        if not opens or not ends:
            continue
        ro, re_ = opens[0], ends[0]
        for c in range(2, vs.max_column):
            # side-by-side FUNDS (Fund 01 | Fund 02 | Fund 04, MDE Attachment A) are not periods: one fund's ending never
            # carries into the next fund's starting. Skip a column whose nearest header above the opening row is "Fund ..".
            hdr = next((vs.cell(row=r, column=c).value.strip() for r in range(ro - 1, 0, -1)
                        if isinstance(vs.cell(row=r, column=c).value, str) and vs.cell(row=r, column=c).value.strip()), "")
            if re.match(r"fund\b", hdr, re.I):
                continue
            e, o = vs.cell(row=re_, column=c).value, vs.cell(row=ro, column=c + 1).value
            if num(e) and num(o):
                rep.did("roll-forward columns")
            if num(e) and num(o) and abs(e - o) > TOL:
                rep.add("HIGH", "rollforward", f"{vs.title}: {vs.cell(row=re_, column=c).coordinate} ending {e:,.2f} does not carry "
                        f"to {vs.cell(row=ro, column=c + 1).coordinate} starting {o:,.2f}")
        tin = [r for r in range(1, vs.max_row + 1) if re.match(r"^total\b.*\bin\b", label_of(vs, r), re.I)
               and not re.search(r"\bout\b", label_of(vs, r), re.I)]
        tout = [r for r in range(1, vs.max_row + 1) if re.match(r"^total\b.*\bout\b", label_of(vs, r), re.I)]
        loc = [r for r in range(1, vs.max_row + 1) if re.search(r"line of credit|\bloc\b", label_of(vs, r), re.I)]
        if tin and tout:
            for c in range(2, vs.max_column + 1):
                vals = [vs.cell(row=r, column=c).value for r in (ro, tin[0], tout[0], re_)]
                if not all(num(x) for x in vals):
                    continue
                other = sum(vs.cell(row=r, column=c).value or 0 for r in loc if num(vs.cell(row=r, column=c).value))
                gap = vals[3] - (vals[0] + vals[1] + vals[2] + other)
                rep.did("start + in + out = end columns")
                if abs(gap) > TOL:
                    rep.add("HIGH", "rollforward", f"{vs.title}: {vs.cell(row=re_, column=c).coordinate} ending {vals[3]:,.2f} is off "
                            f"{gap:,.2f} from starting + total in + total out")


def block_rows(vs):
    """Rows of a CHECKS block: a header like "CHECKS (all must be 0)" (or a bare upper-case CHECKS), then every row down
    to the first blank label. The header's own value is the block total. A CLIENT FLAGS block sits after a blank row on
    purpose (client data findings, not file errors), so it is never part of the block."""
    out = {}
    r = 1
    while r <= vs.max_row:
        lab = label_of(vs, r)
        if CHECKS_HDR.match(lab) or lab == "CHECKS":
            out[r] = lab + " (block total)"
            r += 1
            while r <= vs.max_row and label_of(vs, r) and not FLAG_HDR_RX.match(label_of(vs, r)) and not DIFF_BLOCK_RX.match(label_of(vs, r)):
                out[r] = label_of(vs, r)
                r += 1
        r += 1
    return out


def diff_block_rows(vs):
    """Rows of a DIFFERENCES (explained, not errors) block: the header, then every row down to the first blank label or the next
    Check / CHECKS row. Only a header that sits BELOW a CHECKS header counts; a DIFFERENCES header above CHECKS, or on a sheet with
    no CHECKS block, exempts nothing. A nonzero in the block is the explanation of a check, so it is no finding."""
    first_checks = next((r for r, lab in block_rows(vs).items() if lab.endswith("(block total)")), None)
    out = set()
    if first_checks is None:
        return out
    r = first_checks + 1
    while r <= vs.max_row:
        if r > first_checks and DIFF_BLOCK_RX.match(label_of(vs, r)):
            out.add(r)
            r += 1
            while (r <= vs.max_row and label_of(vs, r) and not CHECK_LABEL.match(label_of(vs, r))
                   and not CHECKS_HDR.match(label_of(vs, r)) and label_of(vs, r) != "CHECKS"
                   and any(isinstance(vs.cell(row=r, column=c).value, str) and vs.cell(row=r, column=c).value.strip() for c in range(3, 6))):
                out.add(r)
                r += 1
            continue
        r += 1
    return out


_DREF = re.compile(r"^=?\s*(?:ABS\()?\s*(?:'([^']+)'|([A-Za-z0-9_]+))!\$?([A-Z]{1,3})\$?(\d+)\s*\)?\s*$", re.I)


def explained_cells(wf):
    """{(sheet, address)} that a Differences-block row of any sheet points at: formula is one cell on a tab (ABS allowed). A
    nonzero Difference/Check-labelled cell at one of these is explained by that row, so it is not a finding."""
    out = set()
    for fs in wf.worksheets:
        for r in diff_block_rows(fs):
            v = fs.cell(row=r, column=2).value
            m = _DREF.match(v) if isinstance(v, str) else None
            if m:
                out.add(((m.group(1) or m.group(2)).lower(), (m.group(3) + m.group(4)).upper()))
    return out


def first_number(vs, r):
    """First numeric cell right of the label: the block's value column, whatever the layout."""
    for c in range(2, vs.max_column + 1):
        v = vs.cell(row=r, column=c).value
        if num(v):
            return c, v
    return None, None


def check_zero_rows(wv, rep, wf=None):
    expl = explained_cells(wf) if wf is not None else set()
    for vs in wv.worksheets:
        blk = block_rows(vs)
        dif = diff_block_rows(vs)
        if dif:
            rep.did("explained differences", len(dif) - 1)    # the header is not one
        for r, lab in blk.items():
            c, v = first_number(vs, r)
            if c is None:
                continue
            rep.did("check cells at zero")
            if abs(v) > TOL:
                rep.add("HIGH", "checks", f"{vs.title}!{vs.cell(row=r, column=c).coordinate} '{lab}' is {v:,.2f}, "
                                          "in a CHECKS block, s/b 0")
        for r in range(1, vs.max_row + 1):
            if r in blk or r in dif:
                continue
            lab = label_of(vs, r)
            if SIGN_NOTE.search(lab):
                continue
            if CHECK_LABEL.match(lab):
                for c in range(2, vs.max_column + 1):
                    v = vs.cell(row=r, column=c).value
                    if num(v) and abs(v) > TOL and (vs.title.lower(), vs.cell(row=r, column=c).coordinate) in expl:
                        rep.did("cells explained by a Differences row")
                        continue
                    if num(v):
                        rep.did("check cells at zero")
                    if num(v) and abs(v) > TOL:
                        rep.add("HIGH", "checks", f"{vs.title}!{vs.cell(row=r, column=c).coordinate} '{lab}' is {v:,.2f}, s/b 0")
            elif DIFF_LABEL.match(lab):
                for c in range(2, vs.max_column + 1):
                    v = vs.cell(row=r, column=c).value
                    if num(v) and abs(v) > TOL and (vs.title.lower(), vs.cell(row=r, column=c).coordinate) in expl:
                        rep.did("cells explained by a Differences row")
                        continue
                    if num(v) and abs(v) > TOL:
                        rep.add("INFO", "checks", f"{vs.title}!{vs.cell(row=r, column=c).coordinate} '{lab}' is {v:,.2f}; a difference line, confirm it is meant to be nonzero")


def check_debit_credit(wv, rep):
    for vs in wv.worksheets:
        for hr in range(1, min(vs.max_row, 15) + 1):
            heads = {str(vs.cell(row=hr, column=c).value or "").strip().lower(): c for c in range(1, vs.max_column + 1)}
            dc, cc = heads.get("debit"), heads.get("credit")
            if not dc or not cc:
                continue
            rep.did("debit/credit sheets")
            grp = next((c for k, c in heads.items() if re.match(r"^(je|je #|batch|entry|journal)", k)), None)
            tot, groups = [0.0, 0.0], {}
            for r in range(hr + 1, vs.max_row + 1):
                if TOTAL_LABEL.match(label_of(vs, r)):
                    continue
                d, c_ = vs.cell(row=r, column=dc).value, vs.cell(row=r, column=cc).value
                d, c_ = (d if num(d) else 0.0), (c_ if num(c_) else 0.0)
                tot[0] += d
                tot[1] += c_
                if grp and vs.cell(row=r, column=grp).value not in (None, ""):
                    g = groups.setdefault(vs.cell(row=r, column=grp).value, [0.0, 0.0])
                    g[0] += d
                    g[1] += c_
            if abs(tot[0] - tot[1]) > TOL:
                rep.add("HIGH", "debits", f"{vs.title}: debits {tot[0]:,.2f} vs credits {tot[1]:,.2f}, off {tot[0] - tot[1]:,.2f}")
            for g, (d, c_) in groups.items():
                if abs(d - c_) > TOL:
                    rep.add("HIGH", "debits", f"{vs.title}: JE {g} debits {d:,.2f} vs credits {c_:,.2f}")
            break


def check_text_numbers_and_dupes(wv, rep):
    for vs in wv.worksheets:
        numeric_cols = {}
        for row in vs.iter_rows():
            for c in row:
                if num(c.value):
                    numeric_cols[c.column] = numeric_cols.get(c.column, 0) + 1
        for row in vs.iter_rows():
            for c in row:
                if isinstance(c.value, str) and numeric_cols.get(c.column, 0) >= 5 and re.fullmatch(r"\s*\(?-?\$?[\d,]+(\.\d+)?\)?\s*", c.value) and any(ch in c.value for ch in ",.$("):
                    rep.add("MED", "text-numbers", f"{vs.title}!{c.coordinate} '{c.value}' is text in a numeric column")
        seen, dupes = {}, 0
        if vs.max_row >= 20:
            for r in range(1, vs.max_row + 1):
                key = tuple(vs.cell(row=r, column=c).value for c in range(1, min(vs.max_column, 8) + 1))
                if sum(1 for k in key if k not in (None, "")) >= 3 and any(num(k) for k in key):
                    if key in seen:
                        dupes += 1
                        if dupes <= 3:
                            rep.add("INFO", "duplicates", f"{vs.title}!{r} repeats row {seen[key]}: {key[:3]}")
                    else:
                        seen[key] = r
            if dupes > 3:
                rep.add("INFO", "duplicates", f"{vs.title}: {dupes} exact duplicate detail rows in all")


def check_voice(wv, rep):
    sys.path.insert(0, str(Path(__file__).parent))
    from ai_tell_scan import hits_in
    for vs in wv.worksheets:
        for row in vs.iter_rows():
            for c in row:
                if isinstance(c.value, str) and (len(c.value.split()) >= 8 or c.row <= 3):
                    tags = hits_in(c.value)
                    if tags:
                        rep.add("MED", "voice", f"{vs.title}!{c.coordinate} {', '.join(tags)}: '{c.value[:60]}'")


def check_fingerprint(path, rep, text=False, sev="MED"):
    """Machine fingerprints (10/26/26): library author stamps, no Office app tag, build talk and git hashes in the text.
    xlsx: MED, reported not blocking, until every generator stamps its author. docx: HIGH, no stamp with a library
    creator. `python tools/doc_fingerprint.py fix <file> --author "<name>"` clears them."""
    sys.path.insert(0, str(Path(__file__).parent))
    import doc_fingerprint as fp
    rep.did("fingerprint")
    for h in fp.metadata_hits(fp.metadata(path)):
        rep.add(sev, "fingerprint", h)
    if text:
        for tag, part, i, snippet in fp.text_hits(path):
            rep.add(sev, "fingerprint", f"{tag}: '{snippet}'")


def check_docx(path, rep):
    """A Word doc before it leaves: fingerprints (metadata and text, HIGH) plus the AI-tell find-list per paragraph (MED)."""
    sys.path.insert(0, str(Path(__file__).parent))
    import doc_fingerprint as fp
    from ai_tell_scan import hits_in
    check_fingerprint(path, rep, text=True, sev="HIGH")
    n = 0
    for part, i, text in fp.paragraphs(path, "word"):
        n += 1
        if len(text.split()) >= 6:
            tags = hits_in(text)
            if tags:
                rep.add("MED", "voice", f"{Path(part).name} para {i} {', '.join(tags)}: '{text.strip()[:60]}'")
    rep.did("paragraphs", n)
    if n == 0:
        rep.add("UNKNOWN", "docx", "no text found in the document body")


def run_lint(path):
    # 180s base plus 120s per MB: a 25K-row GL workbook outran a flat 180s
    limit = min(1800, 180 + int(Path(path).stat().st_size / 1e6 * 120))
    try:
        r = subprocess.run([sys.executable, str(Path(__file__).parent / "wb_lint.py"), path], capture_output=True,
                           encoding="utf-8", timeout=limit)
    except subprocess.TimeoutExpired:
        return f"  [HIGH] lint        wb_lint.py did not finish in {limit}s - not cleared"   # fail closed, never a pass
    return r.stdout.strip() if r.returncode in (0, 1, 2) else f"wb_lint.py exited {r.returncode}: {r.stderr.strip()[:300]}"


def workbook_amounts(wv):
    vals = set()
    for vs in wv.worksheets:
        for row in vs.iter_rows():
            for c in row:
                if num(c.value):
                    vals.add(int(abs(c.value)))          # a page rounds half up or down; accept either dollar
                    vals.add(int(abs(c.value)) + 1)
    return vals


AMOUNT_RX = re.compile(r"(-?)\$\s?([\d,]+(?:\.\d+)?)\s?([kKmM]?)")


def check_html(html_path, xlsx_path, page_text, rep):
    html = Path(html_path).read_text(encoding="utf-8", errors="replace")
    if not re.search(r"<title>[^<]{2,}</title>", html, re.I):
        rep.add("MED", "html", "no <title>")
    if 'name="viewport"' not in html and "name='viewport'" not in html:
        if "<html" in html[:2000].lower():
            rep.add("MED", "html", "no viewport meta - the page will not fit a phone")
        else:  # a fragment: the Artifact wrapper adds width=device-width
            rep.add("INFO", "html", "fragment with no viewport meta; the Artifact wrapper supplies it")
    if "prefers-color-scheme" not in html:
        rep.add("INFO", "html", "no dark-mode block")
    for m in re.finditer(r"\{\{[^}]+\}\}|__[A-Z][A-Z_]{3,}__|\bTODO\b|\bNEED:", html):
        rep.add("HIGH", "html", f"unfilled placeholder '{m.group(0)}'")
    if not page_text:
        rep.add("UNKNOWN", "html-tie", "no --page-text given, so no figure on the page was checked against the workbook. "
                "Save the rendered text from the browser (get_page_text) and pass it in.")
        return
    text = Path(page_text).read_text(encoding="utf-8", errors="replace")
    for tag in ("NaN", "undefined", "Infinity", "$-", "$NaN"):
        if tag in text:
            rep.add("HIGH", "html", f"'{tag}' shows in the rendered page text")
    _, wv = load_both(xlsx_path)
    amounts, rounded = workbook_amounts(wv), 0
    seen = set()
    for m in AMOUNT_RX.finditer(text):
        if m[3]:
            rounded += 1
            continue
        v = int(float(m[2].replace(",", "")) + 0.5)
        if v in seen or v < 100 or v in amounts:
            continue
        seen.add(v)
        ctx = text[max(0, m.start() - 45):m.end() + 25].replace("\n", " ").strip()
        rep.add("MED", "html-tie", f"${v:,} is on the page but not in the workbook: ...{ctx}...")
    if rounded:
        rep.add("INFO", "html-tie", f"{rounded} rounded figure(s) ($k / $M) not tied; check by eye")


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def stamp(mode, path, xlsx, counts):
    LEDGER.parent.mkdir(exist_ok=True)
    rec = {"ts": dt.datetime.now().isoformat(timespec="seconds"), "mode": mode, "file": str(path), "sha256": sha(path),
           "xlsx_sha256": sha(xlsx) if xlsx else None, "checked": counts}
    with open(LEDGER, "a", encoding="utf-8") as f:
        f.write(json.dumps(rec) + "\n")
    snapshot(path, rec["sha256"])


def snapshot(path, digest):
    """Copy a stamped output\\ deliverable to scratch/file-edit-snapshots/<sha>.<ext>, so tools/learning_edge.py scan-files
    can diff Claude's version against a later edit by a person. Workspace only; never fails the stamp."""
    try:
        if os.environ.get("SHIP_CHECK_SNAPSHOTS", "").lower() == "off" or not (_WS / "CLAUDE.md").is_file():
            return
        p = Path(path).resolve()
        if p.suffix.lower() not in (".xlsx", ".xlsm", ".docx", ".pptx") or not str(p).lower().startswith(str((_WS / "output").resolve()).lower() + os.sep):
            return
        dest = Path(os.environ.get("SHIP_CHECK_SNAPSHOT_DIR") or _WS / "scratch" / "file-edit-snapshots") / f"{digest}{p.suffix.lower()}"
        if not dest.exists():
            dest.parent.mkdir(parents=True, exist_ok=True)
            import shutil
            shutil.copy2(p, dest)
    except OSError as e:
        print(f"note: snapshot for the learning edge not saved ({e})", file=sys.stderr)


def stamped(path, xlsx):
    """Newest ledger line for these exact bytes (and the same workbook for an html page), or None."""
    if not LEDGER.exists():
        return None
    want, want_x, hit = sha(path), (sha(xlsx) if xlsx else None), None
    for line in LEDGER.read_text(encoding="utf-8").splitlines():
        try:
            r = json.loads(line)
        except ValueError:
            continue  # a torn line is not a pass
        if r.get("mode") in ("waiver", "validate-data"):
            continue  # a waiver is a recorded decision and validate-data a separate review, never a ship_check pass
        if r.get("sha256") == want and (r.get("xlsx_sha256") == want_x or (want_x is None and r.get("mode") == "html")):
            hit = r
    return hit


def validated(path):
    """Newest data:validate-data record for these exact bytes, or None."""
    if not LEDGER.exists():
        return None
    want, hit = sha(path), None
    for line in LEDGER.read_text(encoding="utf-8").splitlines():
        try:
            r = json.loads(line)
        except ValueError:
            continue
        if r.get("mode") == "validate-data" and r.get("sha256") == want:
            hit = r
    return hit


def still_needs(path, xlsx=None):
    """What these bytes still lack before an outside send: [] when clear."""
    out = [] if stamped(path, xlsx) else ["ship_check"]
    if VALIDATE_REQUIRED and not validated(path):
        out.append("validate-data")
    return out


def cleared(path):
    return not still_needs(path)


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = ap.add_subparsers(dest="mode", required=True)
    px = sub.add_parser("xlsx")
    px.add_argument("file")
    px.add_argument("--voice", action="store_true", help="kept for old calls; voice now runs by default")
    px.add_argument("--no-voice", action="store_true", help="skip the AI-tell scan (10/6/26: on by default)")
    px.add_argument("--no-stamp", action="store_true")
    px.add_argument("--force-reason", default=None, help="Reviewer only: stamp despite a failed xlsx-ship-ready gate; logged to logs/gate.jsonl")
    ph = sub.add_parser("html")
    ph.add_argument("file")
    ph.add_argument("--xlsx", required=True)
    ph.add_argument("--page-text", default=None)
    ph.add_argument("--no-stamp", action="store_true")
    pd_ = sub.add_parser("docx", help="fingerprints and AI tells on a Word doc we built (10/26/26)")
    pd_.add_argument("file")
    pd_.add_argument("--no-stamp", action="store_true")
    pg = sub.add_parser("gate")
    pg.add_argument("--files", nargs="+", required=True)
    pv = sub.add_parser("validated", help="record a data:validate-data pass on these exact bytes (10/31/26)")
    pv.add_argument("file")
    pv.add_argument("--note", required=True, help="what was checked and what it found, short")
    pv.add_argument("--open", type=int, default=0, help="findings left open (flagged in the file, not fixed)")
    ps = sub.add_parser("stamped")
    ps.add_argument("file")
    ps.add_argument("--xlsx", default=None)
    a = ap.parse_args()
    if a.mode == "validated":
        if not Path(a.file).is_file():
            print(f"REFUSED: {a.file} not found")
            sys.exit(2)
        if len(a.note.strip()) < 10:
            print("REFUSED: --note must say what validate-data checked")
            sys.exit(2)
        LEDGER.parent.mkdir(exist_ok=True)
        with open(LEDGER, "a", encoding="utf-8") as f:
            f.write(json.dumps({"ts": dt.datetime.now().isoformat(timespec="seconds"), "mode": "validate-data",
                                "file": str(a.file), "sha256": sha(a.file), "note": a.note.strip(), "open": a.open}) + "\n")
        print(f"VALIDATED  recorded for these exact bytes ({a.open} open)")
        sys.exit(0)
    if a.mode == "gate":
        bad, missing = [], []
        for f in a.files:
            if not Path(f).is_file():
                missing.append(f)
            elif Path(f).suffix.lower() in GATED and still_needs(f):
                bad.append(f"{f} (needs {', '.join(still_needs(f))})")
        if missing:
            print("REFUSED: not found: " + " | ".join(missing))
            sys.exit(2)
        if bad:
            print("NOT STAMPED: " + " | ".join(bad))
            sys.exit(1)
        print("GATE OK: " + ", ".join(Path(f).name for f in a.files if Path(f).suffix.lower() in GATED) or "GATE OK: nothing to gate")
        sys.exit(0)
    if a.mode == "stamped":
        if not Path(a.file).is_file():
            print(f"REFUSED: {a.file} not found")
            sys.exit(2)
        r, need = stamped(a.file, a.xlsx), still_needs(a.file, a.xlsx)
        if not need:
            print(f"STAMPED  {r['ts']}  {r['mode']}  these exact bytes passed ship_check"
                  + ("  + validate-data" if VALIDATE_REQUIRED else ""))
            sys.exit(0)
        if "ship_check" in need:
            print("NOT STAMPED  no pass on record for these exact bytes - run ship_check before it goes out")
        if "validate-data" in need:
            print("NOT VALIDATED  run data:validate-data, then: python tools/ship_check.py validated <file> --note \"...\"")
        sys.exit(1)
    rep = Report()
    target = a.file
    if a.mode == "xlsx":
        wf, wv = load_both(a.file)
        check_cached(wf, wv, rep)
        check_footing(wf, wv, rep)
        check_rollforward(wv, rep)
        check_zero_rows(wv, rep, wf)
        check_debit_credit(wv, rep)
        check_text_numbers_and_dupes(wv, rep)
        if not a.no_voice:
            check_voice(wv, rep)
        check_fingerprint(a.file, rep)
        lint_text = run_lint(a.file)
        for ln in lint_text.splitlines():
            if ln.strip().startswith("[HIGH]"):
                rep.add("HIGH", "lint", ln.strip()[6:].strip())   # a lint HIGH blocks the exit code and the stamp too
    elif a.mode == "docx":
        check_docx(a.file, rep)
    else:
        check_html(a.file, a.xlsx, a.page_text, rep)
    by, notes = rep.render()
    print(f"SHIP CHECK  {a.mode}  {target}")
    print("=" * 78)
    if a.mode in ("xlsx", "docx"):
        print("CHECKED  " + ("; ".join(f"{k} {v}" for k, v in rep.ran.items()) or "nothing - no formulas, totals or roll-forwards found"))
    print(f"\nFINDINGS ({len(by)})" if by else "\nFINDINGS (0) - nothing settled as wrong; see UNKNOWN lines above if any")
    for sev, check, text in by:
        print(f"  [{sev}] {check:12} {text}")
    for n in notes:
        print(f"  {n}")
    if a.mode == "xlsx":
        print("\nWB_LINT (house conventions)")
        print(lint_text)
    code = rep.worst()
    force = getattr(a, "force_reason", None)
    if a.mode == "xlsx" and not a.no_stamp and (code == 0 or force):
        # No stamp without the xlsx-ship-ready gate (recalc, lint, footing)
        sys.path.insert(0, str(Path(__file__).parent))
        import xlsx_gate
        g = xlsx_gate.evaluate(a.file)
        if not g["pass"]:
            print("\nXLSX-SHIP-READY GATE  FAIL")
            for f in g["failures"][:40]:
                print("  " + xlsx_gate.fmt(f))
            if g.get("needs_review"):
                print(f"  NEEDS REVIEW: failed {g['try']} times on these bytes - stop retrying")
            if not force:
                print("\nNOT STAMPED  fix the lines above and rerun")
                sys.exit(1)
        if force:
            extra = [] if code == 0 else [xlsx_gate.F("ship_check", "", f"ship_check exit {code}")]
            xlsx_gate.log_force(a.file, force, g["failures"] + extra)
            stamp("xlsx", a.file, None, dict(rep.ran, forced=force))
            print(f"\nSTAMPED (FORCED)  reason logged to logs/gate.jsonl: {force}")
            sys.exit(0)
        print("\nXLSX-SHIP-READY GATE  PASS")
    if code == 0 and not a.no_stamp:
        stamp(a.mode, a.file, getattr(a, "xlsx", None), rep.ran)
        print("\nSTAMPED  clean pass recorded for this file's exact bytes")
    sys.exit(code)


if __name__ == "__main__":
    main()
