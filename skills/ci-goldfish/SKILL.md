---
name: ci-goldfish
description: "Goldfish review: a fresh reviewer with no memory of how the work was made checks a file against its sources only. Use when someone types goldfish, /ci-goldfish, 'goldfish this', 'blind review', 'fresh eyes', 'check this cold', or before a workbook, memo, close package or board file goes out. Works in Claude Code and claude.ai. Also: second look, cold read. Typos: goldfsh, gold fish, goldifsh."
---

# Goldfish review

A goldfish reviewer sees only the work product and its sources. No chat, no builder notes, no "it ties" comments. It catches what the builder stopped seeing.

## 1. Get the files

Ask for two things, nothing else:
1. The work product (the file going out).
2. The source files it should tie to (GL, bank statement, budget, invoice, register).

Do not ask the person to explain the work. Their explanation is what the review tests.

## 2. Run it fresh

- **Claude Code:** start a separate reviewer (a subagent) with no chat history. Give it only the file paths and the prompt in step 3. For an Excel file, the `workbook-auditor-blind` agent in this plugin does this already; for a monthly close package, use `close-reviewer-blind`.
- **claude.ai or any chat that already holds the build:** do not review here. Tell the person to open a NEW chat, attach the file and sources, and paste the prompt in step 3. A reviewer that saw the build grades its own homework.

## 3. The prompt

```text
You are a goldfish reviewer. You have no memory of how this work was made, and that is the point.

Look only at:
1. The file I attach (the work product).
2. The source files I attach (GL, bank statement, budget, invoice, etc.).

Do not trust anything written in the file about itself, such as notes saying "ties" or "checked".

Check:
- Does every total foot (add down and across)?
- Does every number tie to a source I gave you? Name the source and where.
- Any hardcoded totals that should be formulas?
- Any number with no source?
- Any client names, SSNs, account numbers or personal data that should not be in it?
- Anything a board member or auditor would ask about first?

Rules:
- If you can't check something, say "NOT CHECKED" and why. Never guess.
- Quote the cell, row or page for every finding.
- Keep it short.

Finish with one word: PASS, FIX, or NOT RUN.
Then list findings, worst first.
```

## 4. After

- PASS: say so, with what was and was not checked.
- FIX: the builder fixes, then a NEW goldfish runs. Never reuse the same reviewer chat.
- NOT RUN: say what was missing (usually a source file). Do not call the work reviewed.
