# ci-public

Skills and reviewers for school finance work, free to use.

## Install

No sign-in needed. In Claude Code:

```
/plugin marketplace add scottbrowncpa-creator/ci-public
/plugin install ci-public@ci-public
```

## What is in it

| Item | Kind | What it does |
|---|---|---|
| ci-plain-voice | skill | Rewrites your draft so it does not read like AI. Keeps every number. |
| ci-plainenglish | skill | Explains anything in plain words, no jargon. |
| plan-explainer | agent | Reads a build plan cold and says what is unclear before anyone starts. |
| workbook-auditor-blind | agent | First look at an Excel file with no backstory: footing, ties, formula errors. |
| close-reviewer-blind | agent | First look at a monthly close package with no backstory. |
| plan-redteam | agent | Pre-mortem on a big build plan: the three likeliest ways it fails. |
| ci-protect-data | skill | Masks student IDs and bank numbers in a file before it goes into Claude. |
| ci-wrapup | skill | Saves an end-of-job summary: what was decided, numbers and sources, what is still open. |
| re-read rule | hook | Every prompt, and after a long chat is squeezed to a summary, reminds Claude to re-read files from disk instead of trusting chat memory. Text in `hooks/reread.txt`. |

The two blind auditors run the helper scripts in `tools/` (wb_lint, ship_check, xlsx_probe, close_scan, plus what they call), and a guard hook (`.claude/hooks/blind_guard.py`) that lets them run only those four. They need Python 3 with openpyxl (`pip install openpyxl`; pyxlsb for .xlsb files). Recalc uses Excel on Windows (`tools/excel_recalc.ps1`).

## Re-read rule

Long chats drift: Claude starts working from what it remembers instead of what the file says. The re-read hook adds a short reminder (`hooks/reread.txt`) to every prompt and after a chat is squeezed to a summary. It costs about 150 tokens per prompt. To turn off just this hook and keep the rest of the plugin, set the environment variable `CI_PUBLIC_REREAD=off` (for example in the `env` block of your Claude Code settings). It uses `cat`, which Claude Code on Windows gets through Git Bash.

## Goldfish reviewers

A goldfish reviewer sees only the file and its sources. No chat, no notes from the person who built it. It catches what the builder stopped seeing.

## Credit

Three of the agents (plan-explainer and the two blind auditors) are adapted from the Elephant-Goldfish Model, Copyright 2026 Google LLC, Apache-2.0 (github.com/google/elephant-goldfish-model). See NOTICE.md.

## Setup

ci-wrapup saves notes to your work OneDrive only when `CI_WRAPUP_ONEDRIVE_MATCH` is set to a word in that folder name (for example the firm name). Without it, ci-wrapup hands you the note as a file instead.
