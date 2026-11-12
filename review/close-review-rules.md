# Close review rules

Reference for the `close-reviewer` agent. The Standard Review
Process is the authority; this file records what the machine can settle, what it
cannot, and where the traps are.

## What `tools/close_scan.py` already settles

Run it first, every time. Do not re-derive by hand what it reports.

| Check | Fires when |
|---|---|
| inventory | no FSP, no Standard Review Process workbook, no BS Rec in the folder |
| review Staff tab | an activity line has no initials |
| review Manager / Director tabs | a logged review point has no completion date, initials or response |
| period line | the "For the period ended" line differs between FSP tabs |
| SoFP | Total Assets does not equal Liabilities & Net Assets, per column |
| BvA | a row's variance is neither actual minus budget nor budget minus actual |
| Check Register | $0 entries not marked VOID, duplicate number+amount, gaps inside a number series |
| Aging | conflicting as-of dates in the header, as-of date not the FSP period end, aging total not tying to SoFP Accounts Payable |

It reports UNKNOWN when it cannot evaluate a check. UNKNOWN is not a pass. Open
the tab yourself with `tools/xlsx_probe.py` and settle it, or say plainly that
it is unsettled.

## What the scan cannot do, and you must

1. **SOFP to BS Rec.** The scan only notes that the BS Rec file exists. Opening
   it and tying each balance sheet line is the single most-skipped step in the
   review workbook, and it is a Manager and Director sign-off line at every
   client. Do it.
2. **Accruals.** Payroll (1000/2000), H&W (3000), rent (5600) against the lease
   and deferred rent schedule, management fees (5800). Is each one booked, and
   does the amount match its support?
3. **Balance sheet aging judgment.** AR over 3 months chased, over 6 months
   evaluated for write-off. Accrued liabilities over 6 months challenged.
   Prepaids supported and actually paid. AP suspense account cleared (the account number varies by client; at some clients the suspense number is close to Accounts Payable itself, so check the chart). Every liability-vendor ACH on the bank statement has a matching GL entry. Notes
   payable split current and long term.
4. **Forecast.** Actuals pasted, VLOOKUPs rolled, ending cash tied to the SOCF,
   large variances against the prior forecast explained.
5. **Sensitive vendors** renamed in the check register. Board widgets tie to BvA.

## Traps

- **Month folder numbering is inconsistent by client.** Some number fiscal
  (07 = January at one client), some calendar. List the directory; never
  compute the folder name.
- **BvA column positions are not stable.** The 5/6/7/9 layout in CLAUDE.md holds
  for Sage FSPs, not for everyone. One client runs 2/4/6, another runs two
  side-by-side blocks at 2/3/4 and 6/7/8. The scan maps columns by header text
  and tells you what it found. Never key a number to a column letter you
  assumed.
- **Variance sign flips between revenue and expense sections.** That is correct
  and is not a finding. Only a row that foots neither direction is a finding.
- **A check register can carry two bank accounts on different number series.**
  The gap between series is not a gap. The scan splits on jumps over 500.
- **`.xlsb` exposes cached values only.** No formula text, no fill colors. If a
  question needs formulas, say so rather than guessing.
- **Tab names vary by client and system.** "SoFP" is "Statement of Financial
  Position" at one client and "Gov Balance Sheet" at another. The scan
  resolves by role and prints the mapping. Read the mapping before trusting a
  finding.
- **Two FSPs in one folder** usually means an Original plus a reviewed copy. The
  scan takes the one not named "Original". Say which file you reviewed.

## Reporting

Findings first, severity order, one line each, file and cell reference on every
number. Then what you could not settle, named explicitly. No methodology prose,
no restating the checklist back. If nothing is wrong, say the package is clean
and list what you checked in one line.

Never edit anything in the SharePoint client tree. Write nothing; report.
