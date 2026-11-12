"""Does every formula cell in an xlsx carry a cached value? openpyxl loads it twice; formula error values are listed too.

  python tools/xlsx_cached_check.py <book.xlsx> [<book2.xlsx> ...]

One JSON line per file: path, formulas, uncached, uncached_cells, uncached_more, empty_text, errors,
error_cells, errors_more - or path + error when the file cannot be read.
Exit 0 = every formula cached, no error values. 2 = every formula cached, but some return an error
(#REF!, #DIV/0!, ...). 1 = a formula with no cached value, or a file it could not read (never a pass).
A formula whose result is "" (=IF(x,"",y)) is cached, not missing: openpyxl hands it back as None, but
typed "str" because Excel stored it as text. An openpyxl-written formula carries no type, so an
uncalculated IF(...,"",...) still counts as uncached. Those "" results are counted as empty_text.
Cell lists stop at 10 per file; the counts are always complete and *_more says how many were left off.
Run by tools/excel_recalc.ps1 after every save.
"""
import json
import os
import sys
import warnings

LIST_MAX = 10


def check(path):
    if not os.path.isfile(path):
        return {"path": path, "error": "file not found"}
    if os.path.splitext(path)[1].lower() not in (".xlsx", ".xlsm"):
        return {"path": path, "error": "only .xlsx/.xlsm can be read with openpyxl"}
    try:
        from openpyxl import load_workbook
        wbf = load_workbook(path)                   # formulas
        wbv = load_workbook(path, data_only=True)   # cached values
    except Exception as e:  # corrupt, locked, not a zip: it cannot be verified, so it fails
        return {"path": path, "error": "%s: %s" % (type(e).__name__, e)}
    n = empty = 0
    uncached, errors = [], []
    for ws in wbf.worksheets:
        wv = wbv[ws.title]
        for row in ws.iter_rows():
            for c in row:
                if c.data_type != "f":
                    continue
                n += 1
                v = wv[c.coordinate]
                where = "%s!%s" % (ws.title, c.coordinate)
                if v.data_type == "e":
                    errors.append("%s %s" % (where, v.value))
                elif v.value is None:
                    if v.data_type == "str":
                        empty += 1
                    else:
                        uncached.append(where)
    return {"path": path, "formulas": n,
            "uncached": len(uncached), "uncached_cells": uncached[:LIST_MAX],
            "uncached_more": max(0, len(uncached) - LIST_MAX),
            "empty_text": empty,
            "errors": len(errors), "error_cells": errors[:LIST_MAX],
            "errors_more": max(0, len(errors) - LIST_MAX)}


def main(paths):
    warnings.simplefilter("ignore")  # openpyxl warns about extensions it drops; stdout is JSON only
    if not paths:
        print(json.dumps({"error": "usage: python tools/xlsx_cached_check.py <book.xlsx> [...]"}))
        return 1
    worst = 0
    for p in paths:
        r = check(p)
        print(json.dumps(r), flush=True)
        code = 1 if ("error" in r or r["uncached"]) else (2 if r["errors"] else 0)
        worst = 1 if 1 in (worst, code) else max(worst, code)
    return worst


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
