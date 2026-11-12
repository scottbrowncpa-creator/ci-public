"""Mechanical scan of one monthly close package. Read-only; writes nothing.

Engine behind the close-reviewer agent. It gathers facts and runs the checks a
machine can settle (does it foot, does it tie, is it initialed, is a check
number missing). Judgment stays with the agent reading the output.

    python tools/close_scan.py "<...>/25-26/Financials/12 - June"
    python tools/close_scan.py "<folder>" --json
    python tools/close_scan.py "<folder>" --max-findings 200

Every cap reports itself on a TRUNCATED line. A check that cannot be evaluated
reports UNKNOWN, never silence.
"""

import argparse
import datetime as dt
import json
import os
import re
import sys

MAX_ROWS = 4000          # per sheet
MAX_LIST = 25            # per listed detail (uninitialed lines, gaps, dupes)
TOL = 1.0                # dollars; anything under this is rounding
SERIES_BREAK = 500       # a jump this big means a second check series, not a gap
DATE_RE = re.compile(r"(\d{1,2})/(\d{1,2})/(\d{4})")

# Tab names differ by client and by system. Resolve by role, not by literal name.
TAB_ALIASES = {
    "Cover": ["cover"],
    "BvA": ["bva", "budget vs actual", "budget to actual", "budget vs. actual"],
    "SoFP": ["sofp", "statement of financial position", "balance sheet", "gov balance sheet"],
    "SoCF": ["socf", "statement of cash flows", "cash flow"],
    "Check Register": ["check register", "check/voucher register"],
    "Aging": ["aging", "ap aging", "vendor aging report", "a/p aging"],
}
REVIEW_TABS = ["Staff", "Manager", "Director"]

findings = []
facts = []


def add(sev, area, msg):
    findings.append({"severity": sev, "area": area, "message": msg})


def fact(area, msg):
    facts.append({"area": area, "message": msg})


# ------------------------------------------------------------------ workbook

def ext(p):
    return os.path.splitext(p)[1].lower()


def sheet_names(path):
    if ext(path) == ".xlsb":
        import pyxlsb
        with pyxlsb.open_workbook(path) as wb:
            return list(wb.sheets)
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def grid(path, sheet, max_rows=MAX_ROWS, quiet_cap=False):
    """{(row, col): value} with 1-based indexes. Values only, never formulas."""
    out = {}
    hit_cap = False
    if ext(path) == ".xlsb":
        import pyxlsb
        with pyxlsb.open_workbook(path) as wb:
            with wb.get_sheet(sheet) as ws:
                for i, row in enumerate(ws.rows(), start=1):
                    if i > max_rows:
                        hit_cap = True
                        break
                    for c in row:
                        if c.v is not None and c.v != "":
                            out[(i, c.c + 1)] = c.v
    else:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb[sheet]
            for i, row in enumerate(ws.iter_rows(max_row=max_rows), start=1):
                for c in row:
                    if c.value not in (None, ""):
                        out[(i, c.column)] = c.value
            if ws.max_row and ws.max_row > max_rows:
                hit_cap = True
        finally:
            wb.close()
    if hit_cap and not quiet_cap:
        add("INFO", sheet, "sheet exceeds %d rows; scan covered the first %d only" % (max_rows, max_rows))
    return out


def txt(v):
    return str(v).strip() if v is not None else ""


def num(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.replace(",", "").replace("$", "").strip()
        if s.startswith("(") and s.endswith(")"):
            s = "-" + s[1:-1]
        try:
            return float(s)
        except ValueError:
            return None
    return None


def serial_to_date(v):
    n = num(v)
    if n is None or not (20000 < n < 80000):
        return None
    return dt.date(1899, 12, 30) + dt.timedelta(days=int(n))


def money(x):
    return "${:,.0f}".format(x)


def rows_of(g):
    return sorted({r for r, _ in g})


def row_cells(g, r):
    return {c: v for (rr, c), v in g.items() if rr == r}


def find_text(g, needle, max_row=None, max_col=None):
    """First (row, col) whose text contains needle, case-insensitive."""
    nl = needle.lower()
    for (r, c) in sorted(g):
        if max_row and r > max_row:
            break
        if max_col and c > max_col:
            continue
        v = g[(r, c)]
        if isinstance(v, str) and nl in v.lower():
            return r, c
    return None, None


def resolve_tabs(names):
    """{role: actual sheet name} - exact alias first, then prefix, then contains."""
    out = {}
    lower = {n.lower().strip(): n for n in names}
    for role, aliases in TAB_ALIASES.items():
        hit = None
        for a in aliases:
            if a in lower:
                hit = lower[a]
                break
        if hit is None:
            for a in aliases:
                for ln, orig in lower.items():
                    if ln.startswith(a):
                        hit = orig
                        break
                if hit:
                    break
        if hit is None:
            for a in aliases:
                for ln, orig in lower.items():
                    if a in ln:
                        hit = orig
                        break
                if hit:
                    break
        if hit:
            out[role] = hit
    return out


# ----------------------------------------------------------------- inventory

def classify(name):
    n = name.lower()
    if "standard review process" in n:
        return "review"
    if re.search(r"\bbs\s*rec", n) or "balance sheet rec" in n:
        return "bsrec"
    if "forecast" in n:
        return "forecast"
    if "bank rec" in n:
        return "bankrec"
    if "checklist" in n:
        return "checklist"
    if "fsp" in n:
        return "fsp"
    return "other"


def inventory(folder):
    files = [f for f in sorted(os.listdir(folder))
             if os.path.isfile(os.path.join(folder, f)) and not f.startswith("~$")]
    buckets = {}
    for f in files:
        buckets.setdefault(classify(f), []).append(f)
    for kind, sev, note in (("fsp", "HIGH", ""),
                            ("review", "HIGH", ""),
                            ("bsrec", "MED", " (some clients keep it under Cash Position or Transactions)")):
        if kind not in buckets:
            add(sev, "inventory", "no %s file in this folder%s" % (kind.upper(), note))
    for kind, names in sorted(buckets.items()):
        fact("inventory", "%s: %s" % (kind, ", ".join(names)))
    # Prefer a final FSP over one named "Original"
    fsps = buckets.get("fsp", [])
    final = [f for f in fsps if "original" not in f.lower()]
    if fsps and not final:
        add("MED", "inventory", "only an 'Original' FSP is present; no reviewed copy")
    pick = (final or fsps)
    return {
        "fsp": os.path.join(folder, pick[0]) if pick else None,
        "fsp_all": [os.path.join(folder, f) for f in fsps],
        "bsrec": os.path.join(folder, buckets["bsrec"][0]) if "bsrec" in buckets else None,
        "review": os.path.join(folder, buckets["review"][0]) if "review" in buckets else None,
        "files": files,
    }


# -------------------------------------------------------------- review sheet

BOILERPLATE = (
    "please log any review points here",
    "please indicate any response here",
)


def _checklist_tab(path, tab, g, hdr_r):
    """Staff-style grid: one row per activity, initials in a sign-off column."""
    hdr = row_cells(g, hdr_r)
    sign_c = act_c = None
    for c, v in hdr.items():
        t = txt(v).lower()
        if "activity" in t:
            act_c = c
        if "completed by" in t:
            sign_c = c
    if sign_c is None or act_c is None:
        add("INFO", "review", "%s tab: activity or sign-off column missing; not evaluated" % tab)
        return
    total, blank = 0, []
    for r in rows_of(g):
        if r <= hdr_r:
            continue
        activity = txt(g.get((r, act_c)))
        if len(activity) < 12:          # section headers and spacers
            continue
        total += 1
        if not txt(g.get((r, sign_c))):
            blank.append((r, activity[:70]))
    if total == 0:
        add("INFO", "review", "%s tab: no activity rows detected" % tab)
        return
    if blank:
        add("HIGH", "review", "%s tab: %d of %d lines not initialed" % (tab, len(blank), total))
        for r, a in blank[:MAX_LIST]:
            fact("review", "%s r%d uninitialed: %s" % (tab, r, a))
        if len(blank) > MAX_LIST:
            fact("review", "TRUNCATED: %d more uninitialed %s lines not listed"
                 % (len(blank) - MAX_LIST, tab))
    else:
        fact("review", "%s tab: all %d lines initialed" % (tab, total))


def _notes_tab(path, tab, g, hdr_r):
    """Manager/Director style: a log of review points, each needing a response."""
    hdr = row_cells(g, hdr_r)
    cm = dc = cb = rs = None
    for c, v in hdr.items():
        t = txt(v).lower()
        if t.startswith("comment"):
            cm = c
        elif "date completed" in t:
            dc = c
        elif "completed by" in t:
            cb = c
        elif "response" in t:
            rs = c
    who_r, who_c = find_text(g, "Name:", max_row=6)
    if who_r:
        who = txt(g.get((who_r, who_c + 1)))
        fact("review", "%s tab: reviewer named as %s" % (tab, who or "(blank)"))
        if not who:
            add("MED", "review", "%s tab: reviewer name is blank" % tab)
    if cm is None:
        add("INFO", "review", "%s tab: no Comment column; not evaluated" % tab)
        return
    open_pts, total = [], 0
    for r in rows_of(g):
        if r <= hdr_r:
            continue
        comment = txt(g.get((r, cm)))
        if len(comment) < 12:
            continue
        if any(b in comment.lower() for b in BOILERPLATE):
            continue
        total += 1
        done = txt(g.get((r, dc))) if dc else ""
        by = txt(g.get((r, cb))) if cb else ""
        resp = txt(g.get((r, rs))) if rs else ""
        if not (done or by or resp):
            open_pts.append((r, comment[:70]))
    if total == 0:
        add("MED", "review", "%s tab: no review points logged at all" % tab)
        return
    if open_pts:
        add("HIGH" if tab == "Director" else "MED", "review",
            "%s tab: %d of %d review points have no completion date, initials or response"
            % (tab, len(open_pts), total))
        for r, c in open_pts[:MAX_LIST]:
            fact("review", "%s r%d open: %s" % (tab, r, c))
        if len(open_pts) > MAX_LIST:
            fact("review", "TRUNCATED: %d more open %s points" % (len(open_pts) - MAX_LIST, tab))
    else:
        fact("review", "%s tab: all %d review points closed out" % (tab, total))


def scan_review(path):
    names = sheet_names(path)
    for tab in REVIEW_TABS:
        if tab not in names:
            add("MED", "review", "review workbook has no %s tab" % tab)
            continue
        g = grid(path, tab, max_rows=200)
        hdr_r, _ = find_text(g, "Activity", max_row=20)
        if hdr_r is not None:
            _checklist_tab(path, tab, g, hdr_r)
            continue
        hdr_r, _ = find_text(g, "Date Entered", max_row=20)
        if hdr_r is not None:
            _notes_tab(path, tab, g, hdr_r)
            continue
        add("INFO", "review", "%s tab: layout not recognized; NOT cleared" % tab)


# ---------------------------------------------------------------------- FSP

def period_lines(path, names):
    seen = {}
    for s in names:
        g = grid(path, s, max_rows=8, quiet_cap=True)
        r, c = find_text(g, "for the period ended", max_row=8)
        if r:
            seen[s] = txt(g[(r, c)])
    if not seen:
        add("MED", "fsp", "no 'For the period ended' line found on any tab")
        return None
    vals = set(seen.values())
    if len(vals) > 1:
        add("HIGH", "fsp", "period line differs across tabs: " +
            "; ".join("%s=%r" % (k, v) for k, v in sorted(seen.items())))
    else:
        fact("fsp", "period line consistent across %d tabs: %s" % (len(seen), list(vals)[0]))
    return sorted(vals)[0]


def scan_sofp(path, sheet="SoFP"):
    g = grid(path, sheet, max_rows=300)
    ta_r, _ = find_text(g, "Total Assets")
    ln_r, _ = find_text(g, "Liabilities & Net Assets")
    if ln_r is None:
        ln_r, _ = find_text(g, "Total Liabilities and Net Assets")
    if ta_r is None or ln_r is None:
        add("MED", "SoFP", "UNKNOWN: could not locate Total Assets / Liabilities & Net Assets rows")
        return
    ta = row_cells(g, ta_r)
    ln = row_cells(g, ln_r)
    cols = sorted(set(c for c in ta if num(ta[c]) is not None) &
                  set(c for c in ln if num(ln[c]) is not None))
    if not cols:
        add("MED", "SoFP", "UNKNOWN: no numeric columns shared by the two total rows")
        return
    hdr_r = ta_r
    for r in range(1, ta_r):
        rc = row_cells(g, r)
        if any("balance" in txt(v).lower() or "change" in txt(v).lower() for v in rc.values()):
            hdr_r = r
    for c in cols:
        d = num(ta[c]) - num(ln[c])
        label = txt(g.get((hdr_r, c))) or "col %d" % c
        if abs(d) > TOL:
            add("HIGH", "SoFP", "does not balance in '%s': assets %s vs L&NA %s, off %s"
                % (label, money(num(ta[c])), money(num(ln[c])), money(d)))
        else:
            fact("SoFP", "balances in '%s' at %s" % (label, money(num(ta[c]))))


def scan_bva(path, sheet="BvA"):
    g = grid(path, sheet)
    # Header row = the row in the first 15 that names the most of the three columns.
    best, best_score = None, 0
    for r in range(1, 16):
        joined = " ".join(txt(v).lower() for v in row_cells(g, r).values())
        score = sum(k in joined for k in ("actual", "budget", "variance"))
        if score > best_score:
            best, best_score = r, score
    if best is None or best_score < 2:
        add("MED", "BvA", "UNKNOWN: no header row naming at least two of actual/budget/variance")
        return
    hdr_r = best
    hdr = {c: txt(v) for c, v in row_cells(g, hdr_r).items() if txt(v)}
    fact("BvA", "header row %d: %s" % (hdr_r, ", ".join("%d=%s" % (c, t) for c, t in sorted(hdr.items()))))

    # A BvA often carries several blocks side by side (current month, YTD, annual).
    # Walk left to right; each 'actual' opens a block that its budget and variance join.
    blocks, cur = [], None
    for c in sorted(hdr):
        tl = hdr[c].lower()
        if "%" in tl or "to budget" in tl:
            continue
        if "actual" in tl:
            cur = {"actual": c, "budget": None, "variance": None, "label": None}
            blocks.append(cur)
        elif cur and "budget" in tl and cur["budget"] is None:
            cur["budget"] = c
        elif cur and "variance" in tl and cur["variance"] is None:
            cur["variance"] = c
    blocks = [b for b in blocks if b["budget"] and b["variance"]]
    if not blocks:
        add("MED", "BvA", "UNKNOWN: could not map an actual/budget/variance block from row %d" % hdr_r)
        return
    # Label each block from the banner row above it, if there is one.
    above = row_cells(g, hdr_r - 1)
    for b in blocks:
        cands = [c for c in above if c <= b["actual"] and txt(above[c])]
        b["label"] = txt(above[max(cands)]) if cands else "cols %d/%d/%d" % (b["actual"], b["budget"], b["variance"])

    fact("BvA", "%d block(s): %s" % (len(blocks), "; ".join(
        "%s = %d/%d/%d" % (b["label"], b["actual"], b["budget"], b["variance"]) for b in blocks)))
    first = blocks[0]
    if (first["actual"], first["budget"], first["variance"]) != (5, 6, 7):
        add("INFO", "BvA", "first block is at columns %d/%d/%d, not the usual 5/6/7; any script "
                           "keyed to fixed columns reads the wrong cells here"
            % (first["actual"], first["budget"], first["variance"]))

    for b in blocks:
        rows = []
        for r in rows_of(g):
            if r <= hdr_r:
                continue
            a = num(g.get((r, b["actual"])))
            bu = num(g.get((r, b["budget"])))
            v = num(g.get((r, b["variance"])))
            if a is None or bu is None or v is None:
                continue
            rows.append((r, txt(g.get((r, 1)))[:50], a, bu, v))
        if not rows:
            add("MED", "BvA", "UNKNOWN: block '%s' had no fully populated rows" % b["label"])
            continue
        # A BvA legitimately flips sign between revenue and expense sections
        # (favorable variance), so a row footing either direction is fine. What
        # is not fine is a row that foots neither - that is a stale or typed cell.
        fwd = rev = 0
        bad = []
        for r, lab, a, bu, v in rows:
            ok_f = abs(v - (a - bu)) <= TOL
            ok_r = abs(v - (bu - a)) <= TOL
            if ok_f:
                fwd += 1
            if ok_r:
                rev += 1
            if not (ok_f or ok_r):
                bad.append((r, lab, a, bu, v, a - bu))
        fact("BvA", "block '%s': %d rows, %d foot actual-budget, %d foot budget-actual"
             % (b["label"], len(rows), fwd, rev))
        if bad:
            add("HIGH", "BvA", "block '%s': %d of %d rows where variance is neither "
                               "actual minus budget nor budget minus actual"
                % (b["label"], len(bad), len(rows)))
            for r, lab, a, bu, v, want in bad[:MAX_LIST]:
                fact("BvA", "%s r%d %s: actual %s budget %s variance %s (a-b would be %s)"
                     % (b["label"], r, lab, money(a), money(bu), money(v), money(want)))
            if len(bad) > MAX_LIST:
                fact("BvA", "TRUNCATED: %d more breaks in block '%s'" % (len(bad) - MAX_LIST, b["label"]))


def scan_check_register(path, sheet="Check Register"):
    g = grid(path, sheet)
    hdr_r = no_c = None
    for label in ("Check/Voucher", "Check Number", "Check No", "Voucher No", "Check #", "Document no"):
        hdr_r, no_c = find_text(g, label, max_row=20)
        if hdr_r is not None:
            break
    if hdr_r is None:
        add("MED", "Check Register", "UNKNOWN: no header row found")
        return
    hdr = {c: txt(v) for c, v in row_cells(g, hdr_r).items() if txt(v)}
    amt_c = ven_c = None
    for c, t in hdr.items():
        tl = t.lower()
        if "amount" in tl:
            amt_c = c
        if "vendor" in tl or "payee" in tl:
            ven_c = c
    if amt_c is None:
        add("MED", "Check Register", "UNKNOWN: no Amount column in the header row")
        return

    entries = []
    for r in rows_of(g):
        if r <= hdr_r:
            continue
        a = num(g.get((r, amt_c)))
        key = txt(g.get((r, no_c)))
        if a is None and not key:
            continue
        if a is None:
            continue
        entries.append({"row": r, "no": key, "amt": a, "vendor": txt(g.get((r, ven_c))) if ven_c else ""})
    if not entries:
        add("MED", "Check Register", "UNKNOWN: no check rows read")
        return
    total = sum(e["amt"] for e in entries)
    fact("Check Register", "%d entries totaling %s" % (len(entries), money(total)))

    zeros = [e for e in entries if abs(e["amt"]) < 0.005]
    bad_void = [e for e in zeros if "void" not in (e["vendor"] + e["no"]).lower()]
    if bad_void:
        add("MED", "Check Register", "%d zero-dollar entries not marked VOID" % len(bad_void))
        for e in bad_void[:MAX_LIST]:
            fact("Check Register", "r%d check %s vendor %s is $0 with no VOID" % (e["row"], e["no"], e["vendor"]))
        if len(bad_void) > MAX_LIST:
            fact("Check Register", "TRUNCATED: %d more unmarked $0 entries" % (len(bad_void) - MAX_LIST))

    seen = {}
    dupes = []
    for e in entries:
        k = (e["no"], round(e["amt"], 2))
        if e["no"] and k in seen:
            dupes.append((seen[k], e))
        else:
            seen[k] = e
    if dupes:
        add("MED", "Check Register", "%d duplicate check number and amount pairs" % len(dupes))
        for a, b in dupes[:MAX_LIST]:
            fact("Check Register", "check %s %s appears at r%d and r%d" % (a["no"], money(a["amt"]), a["row"], b["row"]))
        if len(dupes) > MAX_LIST:
            fact("Check Register", "TRUNCATED: %d more duplicates" % (len(dupes) - MAX_LIST))

    seq = sorted({int(e["no"]) for e in entries if e["no"].isdigit() and 100 <= int(e["no"]) <= 999999})
    if len(seq) >= 5:
        # A register can carry two bank accounts on separate number series. Split
        # on any jump over SERIES_BREAK and look for gaps only inside a series,
        # otherwise the space between series reads as thousands of missing checks.
        series, cur = [], [seq[0]]
        for a, b in zip(seq, seq[1:]):
            if b - a > SERIES_BREAK:
                series.append(cur)
                cur = [b]
            else:
                cur.append(b)
        series.append(cur)
        if len(series) > 1:
            fact("Check Register", "%d separate number series detected: %s"
                 % (len(series), ", ".join("%d-%d" % (x[0], x[-1]) for x in series)))
        gaps = []
        for run in series:
            if len(run) < 5:
                fact("Check Register", "series %d-%d has only %d numbers; not tested for gaps"
                     % (run[0], run[-1], len(run)))
                continue
            for a, b in zip(run, run[1:]):
                if b - a > 1:
                    gaps.append((a, b, b - a - 1))
        missing = sum(gg[2] for gg in gaps)
        if gaps:
            add("MED", "Check Register", "%d check numbers missing across %d gaps within a series"
                % (missing, len(gaps)))
            for a, b, n in gaps[:MAX_LIST]:
                fact("Check Register", "gap after %d before %d (%d missing)" % (a, b, n))
            if len(gaps) > MAX_LIST:
                fact("Check Register", "TRUNCATED: %d more gaps" % (len(gaps) - MAX_LIST))
        else:
            fact("Check Register", "no gaps inside any number series")
    else:
        fact("Check Register", "too few purely numeric check numbers to test for gaps")


def scan_aging(path, tabs, period_txt=None):
    g = grid(path, tabs["Aging"])

    # As-of date: the aging must be pulled as of the same date as the FSP, and
    # every as-of stamp in the header block must agree with the others.
    stamps = []
    for (r, c) in sorted(g):
        if r > 10:
            break
        if "as of" not in txt(g[(r, c)]).lower():
            continue
        for c2 in sorted(row_cells(g, r)):
            if c2 <= c:
                continue
            d = serial_to_date(g[(r, c2)])
            if d is None:
                m2 = DATE_RE.search(txt(g[(r, c2)]))
                if m2:
                    d = dt.date(int(m2.group(3)), int(m2.group(1)), int(m2.group(2)))
            if d:
                stamps.append((r, c2, d))
                break
        m = DATE_RE.search(txt(g[(r, c)]))
        if m:
            stamps.append((r, c, dt.date(int(m.group(3)), int(m.group(1)), int(m.group(2)))))
    if len(set(d for _, _, d in stamps)) > 1:
        add("HIGH", "Aging", "header carries conflicting as-of dates: " +
            ", ".join("r%dc%d=%s" % (r, c, d.isoformat()) for r, c, d in stamps))
    if stamps:
        asof = max(d for _, _, d in stamps)
        fact("Aging", "pulled as of %s" % asof.isoformat())
        if period_txt:
            m = re.search(r"([A-Za-z]+)\s+(\d{1,2}),\s*(\d{4})", period_txt)
            if m:
                try:
                    want = dt.datetime.strptime(" ".join(m.groups()), "%B %d %Y").date()
                    if want != asof:
                        add("HIGH", "Aging", "aging is as of %s but the FSP period ends %s"
                            % (asof.isoformat(), want.isoformat()))
                except ValueError:
                    add("INFO", "Aging", "UNKNOWN: FSP period text %r not parseable" % period_txt)
    else:
        add("INFO", "Aging", "UNKNOWN: no as-of date found in the header; staleness not tested")

    tr, _ = find_text(g, "Grand total")
    if tr is None:
        tr, _ = find_text(g, "Total")
    ag_total = None
    if tr:
        for c in sorted(row_cells(g, tr), reverse=True):
            n = num(g[(tr, c)])
            if n is not None:
                ag_total = n
                break
        vals = [num(g[(tr, c)]) for c in sorted(row_cells(g, tr)) if num(g[(tr, c)]) is not None]
        if vals:
            # One cell may be the row's own total; if so use it rather than double count.
            total_cell = None
            for i, v in enumerate(vals):
                rest = sum(vals[:i] + vals[i + 1:])
                if abs(v - rest) <= TOL and len(vals) > 1:
                    total_cell = v
                    break
            ag_total = total_cell if total_cell is not None else sum(vals)
    if ag_total is None:
        add("INFO", "Aging", "UNKNOWN: no total row located; AP tie not tested")
        return
    if "SoFP" not in tabs:
        add("INFO", "Aging", "UNKNOWN: no SoFP tab to tie AP against")
        return
    gs = grid(path, tabs["SoFP"], max_rows=300)
    apr, _ = find_text(gs, "Accounts Payable", max_col=3)
    if apr is None:
        add("INFO", "Aging", "UNKNOWN: no Accounts Payable line on SoFP; AP tie not tested")
        return
    ap = None
    for c in sorted(row_cells(gs, apr)):
        n = num(gs[(apr, c)])
        if n is not None:
            ap = n
            break
    if ap is None:
        add("INFO", "Aging", "UNKNOWN: Accounts Payable row has no numeric value")
        return
    d = ag_total - ap
    if abs(d) > TOL:
        add("HIGH", "Aging", "AP aging total %s does not tie to SoFP Accounts Payable %s, off %s"
            % (money(ag_total), money(ap), money(d)))
    else:
        fact("Aging", "AP aging ties to SoFP at %s" % money(ap))


def scan_fsp(path):
    names = sheet_names(path)
    tabs = resolve_tabs(names)
    fact("fsp", "file: %s" % os.path.basename(path))
    fact("fsp", "tabs: %s" % ", ".join(names))
    fact("fsp", "resolved: " + ", ".join("%s -> %s" % (k, v) for k, v in sorted(tabs.items())))
    for role in TAB_ALIASES:
        if role not in tabs:
            add("MED", "fsp", "no tab resolves to %s" % role)
    period_txt = period_lines(path, [v for v in tabs.values()])
    if "SoFP" in tabs:
        scan_sofp(path, tabs["SoFP"])
    if "BvA" in tabs:
        scan_bva(path, tabs["BvA"])
    if "Check Register" in tabs:
        scan_check_register(path, tabs["Check Register"])
    if "Aging" in tabs:
        scan_aging(path, tabs, period_txt)


# --------------------------------------------------------------------- main

def main():
    p = argparse.ArgumentParser(description="Read-only scan of one close package folder")
    p.add_argument("folder")
    p.add_argument("--json", action="store_true")
    p.add_argument("--max-findings", type=int, default=200)
    args = p.parse_args()

    folder = args.folder
    if not os.path.isdir(folder):
        print("ERROR: not a folder: " + folder)
        sys.exit(2)

    inv = inventory(folder)
    if inv["review"]:
        try:
            scan_review(inv["review"])
        except Exception as e:                       # noqa: BLE001 - report, never swallow
            add("HIGH", "review", "scan failed, NOT cleared: %s: %s" % (type(e).__name__, e))
    if inv["fsp"]:
        try:
            scan_fsp(inv["fsp"])
        except Exception as e:                       # noqa: BLE001
            add("HIGH", "fsp", "scan failed, NOT cleared: %s: %s" % (type(e).__name__, e))
    if inv["bsrec"]:
        fact("bsrec", "present: %s (agent must open it to tie SOFP to the BS Rec)"
             % os.path.basename(inv["bsrec"]))

    order = {"HIGH": 0, "MED": 1, "INFO": 2}
    findings.sort(key=lambda f: order.get(f["severity"], 3))
    shown = findings[:args.max_findings]

    if args.json:
        print(json.dumps({
            "folder": folder,
            "files": inv["files"],
            "findings": shown,
            "facts": facts,
            "truncated": len(findings) > len(shown),
        }, indent=1, default=str))
        return

    print("CLOSE SCAN  %s" % folder)
    print("=" * 78)
    print("\nFINDINGS (%d)" % len(findings))
    if not shown:
        print("  none")
    for f in shown:
        print("  [%-4s] %-15s %s" % (f["severity"], f["area"], f["message"]))
    if len(findings) > len(shown):
        print("  TRUNCATED: %d more findings above --max-findings %d"
              % (len(findings) - len(shown), args.max_findings))
    print("\nFACTS")
    for f in facts:
        print("  %-15s %s" % (f["area"], f["message"]))


if __name__ == "__main__":
    main()
