"""Read-only inspector for Excel workbooks (.xlsx/.xlsm/.xlsb).

The engine behind the close-reviewer and workbook-auditor agents. Never writes.
Every mode that can truncate says so on the last line: TRUNCATED: <what>.

    python tools/xlsx_probe.py sheets  <file>
    python tools/xlsx_probe.py dump    <file> --sheet NAME [--rows 1:80] [--cols A:N] [--formulas]
    python tools/xlsx_probe.py find    <file> --text "Check Figure" [--sheet NAME] [--max 50]
    python tools/xlsx_probe.py row     <file> --sheet NAME --row 12
"""

import argparse
import os
import sys

MAX_CELLS_DEFAULT = 4000


def _ext(path):
    return os.path.splitext(path)[1].lower()


def _die(msg):
    print("ERROR: " + msg)
    sys.exit(2)


def _check(path):
    if not os.path.exists(path):
        _die("file not found: " + path)
    if _ext(path) not in (".xlsx", ".xlsm", ".xlsb"):
        _die("unsupported extension %s (xlsx/xlsm/xlsb only)" % _ext(path))


# ---------------------------------------------------------------- xlsb branch

def _xlsb_rows(path, sheet=None, max_rows=None):
    """Yield (sheet_name, row_index, {col_index: value}) for an .xlsb."""
    import pyxlsb

    with pyxlsb.open_workbook(path) as wb:
        names = wb.sheets
        targets = [sheet] if sheet else names
        for name in targets:
            if name not in names:
                _die("sheet %r not in %s" % (name, names))
            with wb.get_sheet(name) as ws:
                for i, row in enumerate(ws.rows(), start=1):
                    if max_rows and i > max_rows:
                        break
                    cells = {}
                    for c in row:
                        if c.v is not None and c.v != "":
                            cells[c.c + 1] = c.v
                    yield name, i, cells


def _xlsb_sheets(path):
    import pyxlsb

    with pyxlsb.open_workbook(path) as wb:
        return list(wb.sheets)


# ---------------------------------------------------------------- xlsx branch

def _load(path, formulas):
    from openpyxl import load_workbook

    return load_workbook(path, data_only=not formulas, read_only=True)


def col_letter(idx):
    from openpyxl.utils import get_column_letter

    return get_column_letter(idx)


def col_index(letter):
    from openpyxl.utils import column_index_from_string

    return column_index_from_string(letter.upper())


# ---------------------------------------------------------------- range parse

def parse_span(spec, conv):
    if not spec:
        return None, None
    if ":" not in spec:
        v = conv(spec)
        return v, v
    lo, hi = spec.split(":", 1)
    return (conv(lo) if lo else None), (conv(hi) if hi else None)


def fmt(v):
    if v is None:
        return ""
    if isinstance(v, float):
        if v == int(v) and abs(v) < 1e15:
            return str(int(v))
        return "%.2f" % v
    s = str(v).replace("\n", " | ").strip()
    return s[:120]


# ---------------------------------------------------------------- commands

def cmd_sheets(args):
    _check(args.file)
    if _ext(args.file) == ".xlsb":
        for i, n in enumerate(_xlsb_sheets(args.file), 1):
            print("%2d  %s" % (i, n))
        return
    wb = _load(args.file, False)
    for i, ws in enumerate(wb.worksheets, 1):
        print("%2d  %-40s state=%-7s dims=%s" % (i, ws.title, ws.sheet_state, ws.calculate_dimension(force=True)))
    wb.close()


def cmd_dump(args):
    _check(args.file)
    r0, r1 = parse_span(args.rows, int)
    c0, c1 = parse_span(args.cols, col_index)
    r0 = r0 or 1
    r1 = r1 or 200
    c0 = c0 or 1
    c1 = c1 or 30
    budget = args.max_cells
    shown = 0
    truncated = None

    if _ext(args.file) == ".xlsb":
        if args.formulas:
            print("NOTE: xlsb exposes cached values only, no formula text")
        for name, i, cells in _xlsb_rows(args.file, args.sheet, max_rows=r1):
            if i < r0:
                continue
            picked = [(c, v) for c, v in sorted(cells.items()) if c0 <= c <= c1]
            if not picked:
                continue
            if shown + len(picked) > budget:
                truncated = "cell budget %d reached at row %d" % (budget, i)
                break
            shown += len(picked)
            print("r%-4d " % i + " | ".join("%s=%s" % (col_letter(c), fmt(v)) for c, v in picked))
    else:
        wb = _load(args.file, args.formulas)
        names = [args.sheet] if args.sheet else wb.sheetnames
        for name in names:
            if name not in wb.sheetnames:
                _die("sheet %r not in %s" % (name, wb.sheetnames))
            ws = wb[name]
            if len(names) > 1:
                print("=== %s ===" % name)
            for i, row in enumerate(ws.iter_rows(min_row=r0, max_row=r1, min_col=c0, max_col=c1), start=r0):
                picked = [(c.column, c.value) for c in row if c.value not in (None, "")]
                if not picked:
                    continue
                if shown + len(picked) > budget:
                    truncated = "cell budget %d reached at row %d" % (budget, i)
                    break
                shown += len(picked)
                print("r%-4d " % i + " | ".join("%s=%s" % (col_letter(c), fmt(v)) for c, v in picked))
            if truncated:
                break
        wb.close()

    print("-- %d cells shown" % shown)
    if truncated:
        print("TRUNCATED: " + truncated)


def cmd_find(args):
    _check(args.file)
    needle = args.text.lower()
    hits = 0
    truncated = None

    if _ext(args.file) == ".xlsb":
        for name, i, cells in _xlsb_rows(args.file, args.sheet):
            for c, v in sorted(cells.items()):
                if isinstance(v, str) and needle in v.lower():
                    if hits >= args.max:
                        truncated = "--max %d reached" % args.max
                        break
                    hits += 1
                    print("%s!%s%d  %s" % (name, col_letter(c), i, fmt(v)))
            if truncated:
                break
    else:
        wb = _load(args.file, False)
        names = [args.sheet] if args.sheet else wb.sheetnames
        for name in names:
            ws = wb[name]
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and needle in c.value.lower():
                        if hits >= args.max:
                            truncated = "--max %d reached" % args.max
                            break
                        hits += 1
                        print("%s!%s  %s" % (name, c.coordinate, fmt(c.value)))
                if truncated:
                    break
            if truncated:
                break
        wb.close()

    print("-- %d hits" % hits)
    if truncated:
        print("TRUNCATED: " + truncated)


def cmd_row(args):
    args.rows = "%d:%d" % (args.row, args.row)
    args.cols = args.cols or "A:AZ"
    args.formulas = False
    args.max_cells = MAX_CELLS_DEFAULT
    cmd_dump(args)


def main():
    # cp1252 consoles crash on checkbox glyphs (U+2610) in trackers; never die mid-dump.
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass
    p = argparse.ArgumentParser(description="Read-only Excel inspector")
    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("sheets")
    s.add_argument("file")
    s.set_defaults(func=cmd_sheets)

    d = sub.add_parser("dump")
    d.add_argument("file")
    d.add_argument("--sheet")
    d.add_argument("--rows", help="e.g. 1:80")
    d.add_argument("--cols", help="e.g. A:N")
    d.add_argument("--formulas", action="store_true", help="show formula text (xlsx only)")
    d.add_argument("--max-cells", type=int, default=MAX_CELLS_DEFAULT)
    d.set_defaults(func=cmd_dump)

    f = sub.add_parser("find")
    f.add_argument("file")
    f.add_argument("--text", required=True)
    f.add_argument("--sheet")
    f.add_argument("--max", type=int, default=50)
    f.set_defaults(func=cmd_find)

    r = sub.add_parser("row")
    r.add_argument("file")
    r.add_argument("--sheet")
    r.add_argument("--row", type=int, required=True)
    r.add_argument("--cols")
    r.set_defaults(func=cmd_row)

    args = p.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
