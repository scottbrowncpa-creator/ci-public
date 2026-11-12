"""Pre-ship lint for an Excel deliverable against the team's workbook conventions.

Engine behind the workbook-auditor agent. Read-only. Mechanical checks only;
whether a note reads like a CPA wrote it is the agent's call, not this script's.

    python tools/wb_lint.py "output/Client/Audit FY26/File abc 091726.xlsx"
    python tools/wb_lint.py <file> --json
    python tools/wb_lint.py <file> --sheet "Summary"

Conventions enforced here (workspace CLAUDE.md):
  no cell over 40 words, no first person in a cell, no reference to another
  workbook, no hardcoded subtotal, no formula errors, no hidden rows in a
  print-ready tab, Y/N not 1/0 (a 1/0 column that other formulas reference is a
  timing or flag input and is skipped), fills confined to the four sanctioned colors.

Caps report themselves on a TRUNCATED line. A check that cannot be evaluated
prints UNKNOWN rather than passing silently.
"""

import argparse
import json
import os
import re
import sys

MAX_LIST = 20
WORD_CAP = 40
ERRORS = ("#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NULL!", "#NUM!")

# Fills that carry meaning in the team's workbooks, plus brand colors
# that legitimately appear in headers. Keys are the last six hex digits.
SANCTIONED = {
    "FFFF00": "yellow - needs verification or an input",
    "FFFF99": "yellow, light",
    "FFC000": "orange - added account, firm-owned, or PY reimbursable",
    "FFD966": "orange, light",
    "1C424D": "CI teal (header)",
    "7DB03D": "CI green (header)",
    "7EB143": "brand green (header)",
    "F5C449": "CI gold (header)",
    "444755": "CI charcoal (header)",
    "D9D9D9": "grey (header)",
    "F2F2F2": "grey, light (header)",
    "808080": "grey (header)",
    "BFBFBF": "grey (header)",
    "0070C0": "house tab role blue (legend swatch)",
    "92D050": "house tab role green (legend swatch)",
    "A6A6A6": "house tab role grey (legend swatch)",
}


def norm_rgb(rgb):
    """openpyxl hands back 8 hex digits with an alpha or index prefix."""
    if not isinstance(rgb, str) or len(rgb) < 6:
        return None
    return rgb[-6:].upper()

# Singular only. Plural first person shows up verbatim in auditor request text
# ("we need to know...") and flagging it buries the real hits.
# & and / on either side = an abbreviation (P&I, I/O), not a pronoun.
FIRST_PERSON = re.compile(r"(?<![A-Za-z&/+])(I|I'm|I've|I'd|my)(?![A-Za-z'&/+])")   # + too: P+I
# "Title I", "Part I", "Phase I" = a Roman numeral after a label, not a pronoun.
ROMAN_I = re.compile(r"\b(Title|Part|Type|Class|Phase|Schedule|Tier|Article|Section|Chapter|Level|Appendix|Exhibit|Form|Grade|Round|Series|Step|Stage|Category)\s+I\b", re.I)   # TITLE I-A in caps too


def first_person(text):
    """True when a note uses singular first person; label-plus-Roman-numeral I is stripped first."""
    if not re.search(r"[a-z]", text):
        return False   # all caps = system data; "...OF I" is a vendor name cut at INC
    return bool(FIRST_PERSON.search(ROMAN_I.sub(r"\1 x", text)))
# a total row STARTS with its label; "net pay" or "$500K net" in a note is not one
SUBTOTAL_LABEL = re.compile(r"^\s*(total|subtotal|grand total|sum of|net (income|assets|change|position|surplus|deficit|revenue|cash))\b", re.I)

findings = []
facts = []

# Simplest syntax, clear audit trail, no explainer tabs or banners.
README_TAB = re.compile(r"read ?me|instructions|start here|how to|overview|methodology|^about$|^guide$", re.I)
FORMULA_LEN, FORMULA_DEPTH = 200, 5
PROSE_LITERAL = re.compile(r'"([^"]*\s[^"]*\s[^"]*\s[^"]*)"')  # a quoted string of 4+ words
CELL_REF = re.compile(r"\$?[A-Z]{1,3}\$?\d+")


def paren_depth(f):
    d = m = 0
    for ch in re.sub(r'"(?:[^"]|"")*"', '""', f):
        if ch == "(":
            d += 1
            m = max(m, d)
        elif ch == ")":
            d -= 1
    return m


def add(sev, area, msg):
    findings.append({"severity": sev, "area": area, "message": msg})


def fact(area, msg):
    facts.append({"area": area, "message": msg})


# a cell or range reference inside a formula, optionally sheet-qualified; whole-column (D:D) too
_REF = re.compile(r"""(?<![A-Za-z0-9_.'])(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!)?"""
                  r"""(?:\$?([A-Z]{1,3})\$?\d+(?::\$?([A-Z]{1,3})\$?\d+)?|\$?([A-Z]{1,3}):\$?([A-Z]{1,3}))(?![A-Za-z0-9_(])""")


def referenced_columns(wbf):
    """{(sheet lower, col index): {(src sheet lower, src col index), ...}} for every column some formula points at.

    A 1/0 column that other formulas multiply or test is a timing or flag input, not a Yes/No field.
    """
    from openpyxl.utils import column_index_from_string as ci
    refs = {}
    lower = {n.lower(): n for n in wbf.sheetnames}
    for ws in wbf.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not (isinstance(v, str) and v.startswith("=")):
                    continue
                f = re.sub(r'"[^"]*"', '""', v)             # string literals are not references
                for m in _REF.finditer(f):
                    sh = (m.group(1) or m.group(2) or ws.title).replace("''", "'").lower()
                    if sh not in lower:
                        continue
                    a = m.group(3) or m.group(5)
                    b = m.group(4) or m.group(6) or a
                    lo, hi = sorted((ci(a), ci(b)))
                    for col in range(lo, min(hi, lo + 200) + 1):
                        refs.setdefault((sh, col), set()).add((ws.title.lower(), c.column))
    return refs


def words(s):
    return len([w for w in re.split(r"\s+", s.strip()) if w])


def listed(area, sev, headline, items):
    if not items:
        return
    add(sev, area, headline % len(items))
    for it in items[:MAX_LIST]:
        fact(area, it)
    if len(items) > MAX_LIST:
        fact(area, "TRUNCATED: %d more not listed" % (len(items) - MAX_LIST))


def main():
    ap = argparse.ArgumentParser(description="Lint an Excel deliverable before it ships")
    ap.add_argument("file")
    ap.add_argument("--sheet", help="limit to one sheet")
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()

    path = args.file
    if not os.path.exists(path):
        print("ERROR: file not found: " + path)
        sys.exit(2)
    if os.path.splitext(path)[1].lower() not in (".xlsx", ".xlsm"):
        print("ERROR: %s cannot be linted; formulas and formatting are only readable in "
              "xlsx/xlsm. Save a copy and lint that." % os.path.splitext(path)[1])
        sys.exit(2)

    from openpyxl import load_workbook

    wbf = load_workbook(path, data_only=False)      # formulas
    wbv = load_workbook(path, data_only=True)       # cached values

    names = [args.sheet] if args.sheet else wbf.sheetnames
    for n in names:
        if n not in wbf.sheetnames:
            print("ERROR: no sheet %r in %s" % (n, wbf.sheetnames))
            sys.exit(2)

    fact("file", os.path.basename(path))
    fact("file", "%d sheets: %s" % (len(wbf.sheetnames), ", ".join(wbf.sheetnames)))

    # ---- naming convention on the file itself
    # CI_INITIALS: set your own initials; with none set the naming check is skipped.
    ini = os.environ.get("CI_INITIALS") or ""
    if ini and not re.search(r"\b%s\b" % re.escape(ini), os.path.basename(path), re.I):
        add("MED", "naming", "filename has no '%s' initials; house pattern is "
                             "'[Descriptive Name] %s MMDDYY.xlsx'" % (ini, ini))
    if not re.search(r"\b\d{6}\b", os.path.basename(path)):
        add("MED", "naming", "filename has no MMDDYY date stamp")

    hidden_sheets = [ws.title for ws in wbf.worksheets if ws.sheet_state != "visible"]
    if hidden_sheets:
        add("INFO", "structure", "hidden sheets present: %s" % ", ".join(hidden_sheets))

    xrefs = referenced_columns(wbf)
    skipped_binary = []
    long_cells, person_cells, ext_refs, err_cells = [], [], [], []
    hard_subtotals, binary_cells = [], []
    prose_formulas, heavy_formulas, banners = [], {}, []
    odd_colors = {}
    hidden_rows, widths = [], {}
    color_census = {}

    for name in names:
        wsf = wbf[name]
        wsv = wbv[name]

        rd = wsf.row_dimensions
        hid = [r for r, d in rd.items() if d.hidden]
        if hid:
            hidden_rows.append("%s: %d hidden rows (%s%s)"
                               % (name, len(hid), ", ".join(str(x) for x in sorted(hid)[:8]),
                                  "..." if len(hid) > 8 else ""))
        for mr in wsf.merged_cells.ranges:
            if mr.min_row <= 6 and mr.max_col - mr.min_col >= 3:
                banners.append("%s!%s" % (name, mr.coord))
        cd = wsf.column_dimensions
        widths[name] = {k: round(v.width, 1) for k, v in cd.items() if v.width}

        for row in wsf.iter_rows():
            for c in row:
                v = c.value
                coord = "%s!%s" % (name, c.coordinate)

                if isinstance(v, str):
                    if v.startswith("="):
                        # brackets inside a quoted string are text, not a [Book]Sheet! link
                        code = re.sub(r'"(?:[^"]|"")*"', '""', v)
                        if "[" in code and "]" in code:
                            ext_refs.append("%s  %s" % (coord, v[:90]))
                        lits = [m for m in PROSE_LITERAL.findall(v) if len(m.split()) >= 4]
                        if lits and "&" in v:
                            prose_formulas.append("%s  builds text: \"%s\"" % (coord, lits[0][:60]))
                        elif len(v) > FORMULA_LEN or paren_depth(v) > FORMULA_DEPTH:
                            shape = CELL_REF.sub("R", v)  # copied-down formulas count once
                            if shape not in heavy_formulas:
                                heavy_formulas[shape] = "%s  %d chars, %d deep: %s" % (coord, len(v), paren_depth(v), v[:80])
                    else:
                        w = words(v)
                        if w > WORD_CAP:
                            long_cells.append("%s  %d words: %s..." % (coord, w, v[:70]))
                        elif w >= 5 and first_person(v):
                            person_cells.append("%s  %s" % (coord, v[:80]))

                cached = wsv[c.coordinate].value
                if isinstance(cached, str) and cached.strip() in ERRORS:
                    err_cells.append("%s  %s" % (coord, cached.strip()))

                fill = c.fill
                if fill is not None and fill.fill_type == "solid":
                    rgb = norm_rgb(getattr(fill.fgColor, "rgb", None))
                    if rgb and rgb not in ("000000", "FFFFFF"):
                        color_census[rgb] = color_census.get(rgb, 0) + 1
                        if rgb not in SANCTIONED:
                            odd_colors.setdefault(rgb, []).append(coord)

        # ---- hardcoded subtotals: a labeled total row whose number is typed in
        for row in wsf.iter_rows():
            label = None
            for c in row:
                if isinstance(c.value, str) and SUBTOTAL_LABEL.search(c.value) and len(c.value) < 60:
                    label = c.value.strip()
                    break
            if not label:
                continue
            for c in row:
                if isinstance(c.value, (int, float)) and not isinstance(c.value, bool) and c.value != 0:
                    hard_subtotals.append("%s!%s  '%s' = %s typed in, not a formula"
                                          % (name, c.coordinate, label[:40], c.value))

        # ---- 1/0 where the convention is Y/N
        for col in wsf.iter_cols():
            head = None
            vals = []
            for c in col:
                if head is None and isinstance(c.value, str) and c.value.strip():
                    head = c.value.strip()
                    continue
                if c.value is not None:
                    vals.append(c.value)
            if not head or len(vals) < 3:
                continue
            # an all-zero column is dollars with nothing in them, not a flag column
            if all(v in (0, 1) and not isinstance(v, bool) for v in vals) and any(v == 1 for v in vals):
                col_idx = col[0].column
                users = {src for src in xrefs.get((name.lower(), col_idx), ()) if src != (name.lower(), col_idx)}
                if users:
                    skipped_binary.append("%s column under '%s': 1/0 skipped, %d formula column(s) reference it"
                                          % (name, head[:40], len(users)))
                    continue
                binary_cells.append("%s column under '%s': %d cells of 1/0 where the "
                                    "convention is Y/N" % (name, head[:40], len(vals)))

    listed("errors", "HIGH", "%d cells carry a formula error", err_cells)
    listed("links", "HIGH", "%d formulas reference another workbook; deliverables must stand alone", ext_refs)
    listed("subtotals", "HIGH", "%d labeled total rows hold a typed number instead of a formula", hard_subtotals)
    listed("notes", "MED", "%d cells run over 40 words", long_cells)
    listed("formula", "MED", "%d formulas write sentences; put the words in a cell, keep the formula a number or Y/N", prose_formulas)
    listed("formula", "MED", "%d distinct formulas over 200 chars or 5 deep; split into helper columns",
           list(heavy_formulas.values()))
    readme = [n for n in wbf.sheetnames if README_TAB.search(n)]
    if readme:
        add("MED", "structure", "explainer tab(s) %s; files open on the numbers" % ", ".join(readme))
    if banners:
        add("INFO", "banner", "%d merged banner(s) in rows 1-6 (%s); plain caps header reads less built"
            % (len(banners), ", ".join(banners[:5])))
    listed("notes", "MED", "%d cells use first person", person_cells)
    listed("flags", "MED", "%d columns use 1/0 instead of Y/N", binary_cells)
    listed("print", "MED", "%d sheets have hidden rows", hidden_rows)
    if skipped_binary:
        add("INFO", "flagskip", "%d 1/0 columns skipped; other formulas reference them (timing or flag multipliers)"
            % len(skipped_binary))
        for it in skipped_binary[:MAX_LIST]:
            fact("flagskip", it)
    if odd_colors:
        add("INFO", "color", "%d fill colors outside the sanctioned palette, over %d cells"
            % (len(odd_colors), sum(len(v) for v in odd_colors.values())))
        for rgb, cells in sorted(odd_colors.items(), key=lambda kv: -len(kv[1]))[:MAX_LIST]:
            fact("color", "%s on %d cells, first at %s" % (rgb, len(cells), cells[0]))
        if len(odd_colors) > MAX_LIST:
            fact("color", "TRUNCATED: %d more colors" % (len(odd_colors) - MAX_LIST))
    if color_census:
        fact("color", "fill census: " + ", ".join(
            "%s x%d%s" % (k, n, " [%s]" % SANCTIONED[k].split(" ")[0] if k in SANCTIONED else "")
            for k, n in sorted(color_census.items(), key=lambda kv: -kv[1])[:10]))

    # ---- column widths should match across tabs in a board-facing file
    if len(widths) > 1:
        seen = {}
        for sh, w in widths.items():
            for col, val in w.items():
                seen.setdefault(col, {}).setdefault(val, []).append(sh)
        ragged = [c for c, vals in seen.items() if len(vals) > 1 and len(seen[c]) > 1]
        if ragged:
            add("INFO", "print", "column width differs between tabs for %d columns (%s); "
                                 "board packets read better when they match"
                % (len(ragged), ", ".join(sorted(ragged)[:8])))

    if not findings:
        add("INFO", "lint", "no convention breaks found by the mechanical checks")

    order = {"HIGH": 0, "MED": 1, "INFO": 2}
    findings.sort(key=lambda f: order.get(f["severity"], 3))

    if args.json:
        print(json.dumps({"file": path, "findings": findings, "facts": facts}, indent=1, default=str))
        return

    print("WORKBOOK LINT  %s" % path)
    print("=" * 78)
    print("\nFINDINGS (%d)" % len(findings))
    for f in findings:
        print("  [%-4s] %-11s %s" % (f["severity"], f["area"], f["message"]))
    print("\nDETAIL")
    for f in facts:
        print("  %-11s %s" % (f["area"], f["message"]))


if __name__ == "__main__":
    main()
