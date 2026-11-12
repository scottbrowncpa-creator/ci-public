---
name: ci-protect-data
description: "Keep student information and employee bank numbers out of Claude, and warn on Social Security numbers. Use whenever a file or paste may hold SSNs, student names or IDs, birth dates, or bank account or routing numbers - payroll registers, direct deposit lists, benefit bills, W-2/1099 files, SpEd or tuition billing, enrollment rosters, MARSS/ISEE extracts. Also use when anyone says mask, redact, hide the SSNs, or asks if a file is safe to share. Masks with scripts/mask_sensitive.py and reminds about /ci-wrapup. Other words: SSN, student data, banking data, last 4 digits, EE bank numbers, is it safe to share, sensitive data. Typos: datas, sensitve, stuent, redcat."
---

# Protect sensitive data

Most client files are fine in Claude: budgets, GLs, trial balances, FSPs, forecasts. Three things never go in unmasked. SSNs get a warning, not a mask: say once that the file has them and carry on.

| Never in unmasked | Masked to |
|---|---|
| Student names, student IDs, MARSS/state IDs, birth dates | Student 001, Student 002 (same kid, same tag) |
| Employee bank account and routing numbers | ****1234 (last 4) |
| Bank or system logins and passwords | never, not even masked |

Counts and totals of students are fine. Employee names and pay amounts are fine.

## Before a file goes in
If the file looks like a payroll register, direct deposit list, benefit bill, tax form, student roster, or SpEd/tuition billing:
1. Run `python scripts/mask_sensitive.py "<file>" --check` to see what it would mask.
2. Run `python scripts/mask_sensitive.py "<file>"` to write `<file> MASKED.xlsx` next to it. The original is never touched.
3. Use the MASKED copy. Glance at it first - the tool finds columns by their header names, so a roster with plain "First Name / Last Name" headers and no "Student" in the title will NOT be caught. Rename the header or mask by hand.

On the Claude website (no scripts), tell the person to delete or blank those columns in Excel before uploading.

## If sensitive data shows up anyway
- Say so plainly, once: "This has SSNs / student names / bank numbers in it."
- Do not repeat the numbers or names back in the reply. Use last 4 or the tag.
- Do the work on the masked version.
- At the end of the chat, remind once:
  > "This chat had sensitive data. Type /ci-wrapup to save a summary to your work OneDrive (Claude Wrap-ups), then delete this chat."
- If the person types "ci-wrapup off", stop reminding for the rest of that chat.

## Account settings everyone should have
- Settings > Privacy > "Help improve our AI models" = OFF. Off keeps chats out of training and on the 30-day retention, not five years.
