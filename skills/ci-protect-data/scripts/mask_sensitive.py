"""Mask student identifiers and bank account/routing numbers in an xlsx or csv before it goes into Claude. SSNs are warned about, not masked.

    python mask_sensitive.py <file>            write "<name> MASKED.<ext>" next to it, print what changed
    python mask_sensitive.py <file> --check    list what it would mask, write nothing (exit 1 if anything found)

Never touches the original. Exit codes: 0 clean or masked, 1 --check found items, 2 refused (unreadable,
unsupported type, or the MASKED copy already exists and is open/locked).

What it warns about (left as is):
  - SSN-shaped text anywhere (123-45-6789 / 123 45 6789), or any value in an SSN / social security column
What it masks:
  - Any value in a column whose header names a bank account or routing #  -> ****last4
  - Student name / student ID / date of birth columns                      -> Student 001, 002... (same kid = same tag)
Leaves alone: employee names, vendor EINs (12-3456789), dollar amounts, phone numbers, plain 9-digit numbers
outside a flagged column.
"""
import csv
import os
import re
import sys

SSN_TEXT = re.compile(r"(?<![\d-])(\d{3})([- ])(\d{2})\2(\d{4})(?![\d-])")
H_SSN = re.compile(r"\bssn\b|social\s*sec|\bss\s*#|\bsocial\b", re.I)
H_BANK = re.compile(r"routing|\baba\b|(bank|acct|account|checking|savings|dd)\s*(no|num|number|#)|account\s*#|acct\s*#|bank\s*acc", re.I)
# Student identifiers only - "Students" or "Student count" (a number of kids) is not flagged.
H_STUDENT = re.compile(r"(student|pupil|\bstu)\s*(name|id\b|#|no\b|num|number|first|last)|marss|state\s*(student\s*)?id|"
                       r"\bdob\b|date\s*of\s*birth|birth\s*date|guardian", re.I)
H_NOT_BANK = re.compile(r"\bgl\b|ledger|account\s*(name|desc|title|code)|object|fund|function", re.I)


def last4(v):
    d = re.sub(r"\D", "", str(v))
    return d[-4:] if len(d) >= 4 else ""


def classify(header):
    h = str(header or "").strip()
    if not h:
        return None
    if H_SSN.search(h):
        return "ssn"
    if H_STUDENT.search(h):
        return "student"
    if H_BANK.search(h) and not H_NOT_BANK.search(h):
        return "bank"
    return None


class Masker:
    def __init__(self):
        self.tags, self.log, self.warn = {}, [], []

    def student(self, v):
        k = str(v).strip().lower()
        if k not in self.tags:
            self.tags[k] = f"Student {len(self.tags) + 1:03d}"
        return self.tags[k]

    def cell(self, v, kind, where):
        if v is None or (isinstance(v, str) and not v.strip()):
            return v
        if isinstance(v, str) and v.startswith("="):
            return v  # formula: leave it, it points at cells we mask
        if kind == "ssn" or (kind is None and isinstance(v, str) and SSN_TEXT.search(v)):
            self.warn.append((where, "ssn"))  # warn only, never masked (10/14/26)
            return v
        if kind == "bank":
            new = "****" + last4(v) if last4(v) else "****"
        elif kind == "student":
            new = self.student(v)
        else:
            return v
        if new != v:
            self.log.append((where, kind))
        return new


def header_row(rows):
    """First row (of the first 10) with at least two text cells is the header."""
    for i, r in enumerate(rows[:10]):
        if sum(1 for c in r if isinstance(c, str) and c.strip()) >= 2:
            return i
    return None


def do_csv(path, out, check):
    m = Masker()
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    hi = header_row(rows)
    kinds = [classify(h) for h in rows[hi]] if hi is not None else []
    new = []
    for ri, r in enumerate(rows):
        nr = []
        for ci, v in enumerate(r):
            kind = kinds[ci] if hi is not None and ri > hi and ci < len(kinds) else None
            nr.append(m.cell(v, kind, f"row {ri + 1} col {ci + 1}"))
        new.append(nr)
    if not check:
        with open(out, "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerows(new)
    return m


def do_xlsx(path, out, check):
    import openpyxl
    m = Masker()
    wb = openpyxl.load_workbook(path)
    for ws in wb.worksheets:
        rows = [[c.value for c in r] for r in ws.iter_rows(max_row=10)]
        hi = header_row(rows)
        kinds = [classify(h) for h in rows[hi]] if hi is not None else []
        for r in ws.iter_rows():
            for c in r:
                kind = kinds[c.column - 1] if hi is not None and c.row > hi + 1 and c.column - 1 < len(kinds) else None
                nv = m.cell(c.value, kind, f"{ws.title}!{c.coordinate}")
                if nv is not c.value and not check:
                    c.value = nv
    if not check:
        wb.save(out)
    return m


def main(argv):
    args = [a for a in argv if not a.startswith("--")]
    check = "--check" in argv
    if len(args) != 1:
        print(__doc__)
        return 2
    path = args[0]
    base, ext = os.path.splitext(path)
    ext = ext.lower()
    if ext not in (".xlsx", ".xlsm", ".csv"):
        print(f"REFUSED: {ext or 'no extension'} not supported. Save as .xlsx or .csv first.")
        return 2
    out = f"{base} MASKED{ext}"
    try:
        m = do_csv(path, out, check) if ext == ".csv" else do_xlsx(path, out, check)
    except Exception as e:  # unreadable = not cleared, never "nothing found"
        print(f"REFUSED: could not read or write {os.path.basename(path)} ({type(e).__name__}: {e})")
        return 2
    if m.warn:
        print(f"WARNING: {len(m.warn)} SSN-like value(s) left as is, first at {m.warn[0][0]}")
    counts = {}
    for _, k in m.log:
        counts[k] = counts.get(k, 0) + 1
    if not m.log:
        print("Nothing to mask found." + ("" if check else f" Copy written anyway: {os.path.basename(out)}"))
        return 0
    print(("WOULD MASK" if check else "MASKED") + ": " + ", ".join(f"{v} {k}" for k, v in sorted(counts.items())))
    for where, k in m.log[:25]:
        print(f"  {where}: {k}")
    if len(m.log) > 25:
        print(f"  ...and {len(m.log) - 25} more (list capped at 25; all were masked)")
    if not check:
        print(f"Saved: {out}\nUse the MASKED copy in Claude. Original untouched.")
    return 1 if check else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
