"""xlsx-ship-ready gate: recalc cached values, wb_lint errors, footing and check figures - an xlsx must pass before ship_check will stamp it.

    python tools/xlsx_gate.py check "<file.xlsx>"                 gate the file as it is (no Excel)
    python tools/xlsx_gate.py check "<file.xlsx>" --recalc        run tools/excel_recalc.ps1 on it first (temp-copy safe)
    python tools/xlsx_gate.py check "<file.xlsx>" --copy-to DIR   dry run: copy to DIR, recalc the copy, gate the copy
    add --dry-run to log the verdict without counting a retry, --json for machine output.

Steps, in order; the first step that fails stops the run (later steps would judge stale numbers):
  1 recalc   every formula has a cached value (tools/xlsx_cached_check.py; --recalc runs excel_recalc.ps1 first)
  2 lint     tools/wb_lint.py: formula errors, external refs, typed subtotals, cells over 40 words, 1/0 columns
  3 footing  every Total/Subtotal row whose formula only adds cells in its own column adds its whole block: no
             numbered detail row (or, for a grand total, no subtotal) left out between the first and last row it
             adds. A deliberate partial range (Total cash in = inflows only) passes. Check / Difference /
             Variance to rows and CHECKS blocks are 0.
Each failure prints  STEP  Sheet!Cell  what is wrong  expected X  actual Y  - an agent can fix it from the line.

Retry contract: every verdict goes to logs/gate.jsonl (gate xlsx-ship-ready, sha256). The 3rd failure on the same
file hash (dry runs not counted) lands the file in logs/xlsx-needs-review.json; stop retrying and tell the reviewer.
Exit 0 pass, 1 fail. Any exception is a fail. Tests: python tools/test_xlsx_gate.py
"""

import argparse
import datetime as dt
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import warnings
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))
GATE = "xlsx-ship-ready"
# Outside the workspace (a staff plugin install) ROOT is the skill folder, wiped on every plugin update,
# so logs live in the person's home folder instead.
_LOGS = (ROOT / "logs") if (ROOT / "CLAUDE.md").is_file() else (Path.home() / ".ci-tools")
LOG = Path(os.environ.get("XLSX_GATE_LOG") or _LOGS / "gate.jsonl")
NEEDS = Path(os.environ.get("XLSX_GATE_NEEDS") or _LOGS / "xlsx-needs-review.json")
MAX_TRIES = 3
LINT_FAIL_AREAS = {"errors", "links", "subtotals", "flags"}   # plus notes, but only the 40-word half

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def F(step, cell, what, expected="", actual=""):
    return {"step": step, "cell": cell, "what": what, "expected": expected, "actual": actual}


def fmt(f):
    s = f"{f['step']:8} {f['cell']}  {f['what']}"
    if f["expected"] != "" or f["actual"] != "":
        s += f"  expected {f['expected']}  actual {f['actual']}"
    return s


# ------------------------------------------------------------------ step 1
def run_recalc(path):
    """excel_recalc.ps1 works on a temp copy and writes back only when every formula is cached."""
    r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(HERE / "excel_recalc.ps1"),
                        "-Paths", str(path)], capture_output=True, encoding="utf-8", errors="replace", timeout=600)
    return r.returncode, (r.stdout + r.stderr).strip()


def step_recalc(path):
    import xlsx_cached_check
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        r = xlsx_cached_check.check(str(path))
    if "error" in r:
        return [F("recalc", Path(path).name, f"cannot read workbook: {r['error']}")]
    out = []
    for c in r["uncached_cells"]:
        out.append(F("recalc", c, "formula has no cached value - run tools/excel_recalc.ps1 -Paths <file>",
                     "a calculated value", "empty"))
    if r["uncached_more"]:
        out.append(F("recalc", "(more)", f"{r['uncached_more']} more uncached formula cells not listed"))
    return out


# ------------------------------------------------------------------ step 2
def step_lint(path):
    p = subprocess.run([sys.executable, str(HERE / "wb_lint.py"), str(path), "--json"], capture_output=True,
                       encoding="utf-8", errors="replace", timeout=300)
    try:
        d = json.loads(p.stdout)
    except ValueError:
        return [F("lint", Path(path).name, f"wb_lint.py gave no JSON (exit {p.returncode}): {p.stderr.strip()[:200]}")]
    out = []
    what = {"errors": "formula returns an error", "links": "formula points at another workbook",
            "subtotals": "total row holds a typed number; make it a formula",
            "flags": "1/0 column; use Y/N", "notes": "cell over 40 words; cut to a short note"}
    fail_areas = {f["area"] for f in d.get("findings", [])
                  if f["area"] in LINT_FAIL_AREAS or (f["area"] == "notes" and "40 words" in f["message"])}
    for fa in d.get("facts", []):
        a, m = fa["area"], fa["message"]
        if a not in fail_areas:
            continue
        if a == "notes" and not re.search(r"\d+ words:", m):
            continue                                   # first-person lines share the area; not a gate item
        if m.startswith("TRUNCATED"):
            out.append(F("lint", "(more)", m))
            continue
        cell, _, rest = m.partition("  ")
        if a == "errors":
            out.append(F("lint", cell, what[a], "a value", rest.strip()))
        elif a == "subtotals":
            out.append(F("lint", cell, what[a], "=SUM(...) formula", rest.strip()))
        elif a == "flags":
            out.append(F("lint", m.split(" column")[0] + " (column)", what[a] + ": " + m, "Y/N", "1/0"))
        elif a == "notes":
            out.append(F("lint", cell, what[a], "<= 40 words", rest.split(":")[0].strip()))
        else:
            out.append(F("lint", cell, what[a], "no external reference", rest.strip()[:90]))
    return out


# ------------------------------------------------------------------ step 3
REF_RX = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)(?:\s*:\s*\$?([A-Z]{1,3})\$?(\d+))?")


def additive_same_col(formula, col_letters):
    """Rows the formula adds, if it only adds cells of its own column (=SUM(B2:B9), =B2+B5, =SUM(B3,B8)); else None."""
    body = formula.lstrip("=").replace(" ", "").upper()
    body = re.sub(r"^SUBTOTAL\(9,", "SUM(", body)
    if re.sub(r"SUM\(|\)|,|\+|\$?[A-Z]{1,3}\$?\d+(:\$?[A-Z]{1,3}\$?\d+)?", "", body):
        return None                                    # anything else in there: not a plain add
    rows = []
    for m in REF_RX.finditer(body):
        c1, r1, c2, r2 = m[1], int(m[2]), m[3] or m[1], int(m[4] or m[2])
        if c1 != col_letters or c2 != col_letters:
            return None
        rows.extend(range(min(r1, r2), max(r1, r2) + 1))
    return rows or None


def label(ws, row):
    """First text cell in columns A-D that is a name, not an account code like '431100' or '10-100-1000'."""
    for c in range(1, 5):
        v = ws.cell(row=row, column=c).value
        if isinstance(v, str) and v.strip() and not v.startswith("=") and not re.fullmatch(r"[\d\s.\-/]+", v.strip()):
            return v.strip()
    return ""


SIDE_RX = {"rev": re.compile(r"\b(revenues?|receipts|income|inflows?)\b", re.I),
           "exp": re.compile(r"\b(expenses?|expenditures?|disbursements|outflows?)\b", re.I)}


def opposite_side(total_label, row_label):
    """True when a 'Total revenue' row sits beside a 'Total expense' row (or the reverse): a revenue total never adds expense totals."""
    def side(t):
        r, e = bool(SIDE_RX["rev"].search(t)), bool(SIDE_RX["exp"].search(t))
        return "rev" if r and not e else "exp" if e and not r else None
    a, b = side(total_label), side(row_label)
    return a is not None and b is not None and a != b


def step_footing(path):
    import openpyxl
    from openpyxl.utils import get_column_letter
    import ship_check as sc
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wf = openpyxl.load_workbook(path)
        wv = openpyxl.load_workbook(path, data_only=True)
    out, ran = [], 0
    for ws in wf.worksheets:
        vs = wv[ws.title]
        is_total = {r: bool(sc.TOTAL_LABEL.match(label(ws, r))) for r in range(1, ws.max_row + 1)}
        for r in range(1, ws.max_row + 1):
            if not is_total[r]:
                continue
            lab = label(ws, r)
            for c in range(2, ws.max_column + 1):
                f, v = ws.cell(row=r, column=c).value, vs.cell(row=r, column=c).value
                if not (isinstance(f, str) and f.startswith("=") and sc.num(v)):
                    continue
                L = get_column_letter(c)
                refs = additive_same_col(f, L)
                if refs is None:
                    continue                           # nets, ratios, cross-column math: not a footing row
                refs = set(refs)
                isf = lambda x: isinstance(ws.cell(row=x, column=c).value, str) and str(ws.cell(row=x, column=c).value).startswith("=")
                rollup = all(isf(x) for x in refs)        # adds formula lines (subtotals by any label): gap = another formula line
                kind_total = any(is_total.get(x) for x in refs)   # adds any total: only other totals can be skipped
                # interior gap: a numbered row of the same kind between the first and last row added, left out
                skipped = [x for x in range(min(refs), max(refs) + 1) if x not in refs and x != r
                           and sc.num(vs.cell(row=x, column=c).value) and bool(is_total.get(x)) == kind_total
                           and (not rollup or isf(x))
                           and not re.match(r"^\s*grand", label(ws, x), re.I)
                           and not opposite_side(lab, label(ws, x))]
                covered, todo = set(), list(refs)
                while todo and len(covered) < 5000:
                    x = todo.pop()
                    if x in covered:
                        continue
                    covered.add(x)
                    fx = ws.cell(row=x, column=c).value
                    if isinstance(fx, str) and fx.startswith("="):
                        todo.extend(additive_same_col(fx, L) or [])
                skipped = [x for x in skipped if x not in covered]
                ran += 1
                if skipped:
                    miss = sum(vs.cell(row=x, column=c).value for x in skipped)
                    if abs(miss) > sc.TOL:
                        rows = ", ".join(f"{L}{x}" for x in skipped[:6])
                        out.append(F("footing", f"{ws.title}!{L}{r}",
                                     f"'{lab}' skips {rows} inside its block ({'subtotals' if kind_total else 'detail'})",
                                     f"{v + miss:,.2f}", f"{v:,.2f}"))
        # check figures
        blk = sc.block_rows(vs)
        for r, lab in blk.items():
            c, v = sc.first_number(vs, r)
            if c is not None:
                ran += 1
                if abs(v) > sc.TOL:
                    out.append(F("checks", f"{vs.title}!{get_column_letter(c)}{r}", f"'{lab}' in a CHECKS block", "0", f"{v:,.2f}"))
        for r in range(1, vs.max_row + 1):
            if r in blk:
                continue
            lab = label(vs, r)
            if not lab or sc.SIGN_NOTE.search(lab) or re.match(r"^\s*check\s*(#|no\b|num|register|date|amount)", lab, re.I):
                continue
            is_diff = not sc.CHECK_LABEL.match(lab) and sc.DIFF_LABEL.match(lab)
            if sc.CHECK_LABEL.match(lab) or is_diff:
                for c in range(2, vs.max_column + 1):
                    v = vs.cell(row=r, column=c).value
                    fv = ws.cell(row=r, column=c).value
                    if is_diff and not (isinstance(fv, str) and fv.startswith("=")):
                        continue                       # a typed Difference is a stated amount, not a computed check
                    if sc.num(v):
                        ran += 1
                        if abs(v) > sc.TOL:
                            out.append(F("checks", f"{vs.title}!{get_column_letter(c)}{r}", f"'{lab}' is a check figure", "0", f"{v:,.2f}"))
    return out, ran


# ------------------------------------------------------------------ verdict, retries
def tries_for(h):
    n = 0
    if LOG.exists():
        for line in LOG.read_text(encoding="utf-8", errors="replace").splitlines():
            try:
                r = json.loads(line)
            except ValueError:
                continue
            if r.get("gate") == GATE and r.get("sha256") == h and not r.get("pass") and not r.get("dry_run") and not r.get("forced"):
                n += 1
    return n


def mark_needs_review(path, h, failures):
    cur = []
    if NEEDS.exists():
        try:
            cur = json.loads(NEEDS.read_text(encoding="utf-8"))
        except ValueError:
            cur = []
    if any(c.get("sha256") == h for c in cur):
        return
    cur.append({"file": str(path), "sha256": h, "ts": dt.datetime.now().isoformat(timespec="seconds"),
                "failures": [fmt(f) for f in failures[:15]]})
    NEEDS.parent.mkdir(parents=True, exist_ok=True)
    NEEDS.write_text(json.dumps(cur, indent=1, ensure_ascii=False), encoding="utf-8")


def log(rec):
    LOG.parent.mkdir(parents=True, exist_ok=True)
    rec.setdefault("ts", dt.datetime.now().isoformat(timespec="seconds"))
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(json.dumps(rec, ensure_ascii=False) + "\n")


def evaluate(path, recalc=False, copy_to=None, dry_run=False):
    """Run the three steps. Returns the logged record: pass, failures, sha256, try, needs_review."""
    src = Path(path)
    rec = {"gate": GATE, "file": str(src), "pass": False, "failures": [], "dry_run": bool(dry_run or copy_to)}
    try:
        if not src.is_file() or src.suffix.lower() not in (".xlsx", ".xlsm"):
            raise ValueError("not an .xlsx/.xlsm file on disk")
        target = src
        if copy_to:
            Path(copy_to).mkdir(parents=True, exist_ok=True)
            target = Path(copy_to) / src.name
            shutil.copy2(src, target)
            rec["checked_copy"] = str(target)
        if recalc or copy_to:
            code, msg = run_recalc(target)
            rec["recalc"] = {"exit": code, "msg": msg[-300:]}
        rec["sha256"] = sha(src if not copy_to else target)
        fails = step_recalc(target)
        if not fails:
            fails = step_lint(target)
        if not fails:
            fails, ran = step_footing(target)
            rec["footing_cells_checked"] = ran
        rec["failures"] = fails
        rec["pass"] = not fails
        rec["step_failed"] = fails[0]["step"] if fails else None
    except BaseException as e:   # fail closed
        rec["pass"] = False
        rec["failures"] = rec["failures"] + [F("error", str(src), f"gate could not evaluate: {type(e).__name__}: {e}")]
        rec.setdefault("sha256", sha(src) if src.is_file() else None)
    if not rec["pass"] and not rec["dry_run"] and rec.get("sha256"):
        rec["try"] = tries_for(rec["sha256"]) + 1
        if rec["try"] >= MAX_TRIES:
            rec["needs_review"] = True
            mark_needs_review(src, rec["sha256"], rec["failures"])
    try:
        log(dict(rec))
    except Exception as e:
        rec["pass"] = False
        rec["failures"].append(F("error", str(LOG), f"could not write gate log: {e}"))
    return rec


def log_force(path, reason, failures):
    log({"gate": GATE, "file": str(path), "sha256": sha(path), "pass": False, "forced": True, "reason": reason,
         "failures": failures})


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = ap.add_subparsers(dest="cmd", required=True)
    c = sub.add_parser("check")
    c.add_argument("file")
    c.add_argument("--recalc", action="store_true")
    c.add_argument("--copy-to", default=None)
    c.add_argument("--dry-run", action="store_true")
    c.add_argument("--json", action="store_true")
    a = ap.parse_args()
    rec = evaluate(a.file, recalc=a.recalc, copy_to=a.copy_to, dry_run=a.dry_run)
    if a.json:
        print(json.dumps(rec, ensure_ascii=False, indent=1))
    else:
        print(f"XLSX GATE  {'PASS' if rec['pass'] else 'FAIL'}  {a.file}")
        for f in rec["failures"][:40]:
            print("  " + fmt(f))
        if len(rec["failures"]) > 40:
            print(f"  TRUNCATED: showing 40 of {len(rec['failures'])}")
        if rec.get("needs_review"):
            print(f"  NEEDS REVIEW: failed {rec['try']} times on these bytes - stop retrying ({NEEDS})")
    return 0 if rec["pass"] else 1


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except BaseException as e:
        print(f"XLSX GATE  FAIL  gate error: {e}")
        sys.exit(1)
