# Privacy & data-handling rules — read before adding anything

This repo holds a **planning tool**, not anyone's data. One rule keeps it publishable
and keeps every user's specifics (financial, medical, credentials) out of harm's way.

## The one rule

> **The repo holds the TOOL. It never holds anyone's DATA.**

- **Tool** = the engine, the wizard, the template, tests, and **dummy** sample data
  (fictional people, fictional accounts). ✅ Belongs here.
- **Data** = a real household's filled workbook and everything in it. ❌ Never here.
  It lives outside the repo, on the machine of the person it belongs to.

## Never put these in the repo

- A real filled workbook, or any export of one
- Bank / account numbers, sort codes, balances, policy references
- Passwords, PINs, card numbers, security answers, OAuth credential files
- NHS numbers, medication lists, diagnoses, or other health specifics
- Real names, addresses, or dates of birth of real people — sample data is
  **fictional people only**
- Calendar IDs, tokens, or `google.json` config contents

If you're unsure whether something is "data," treat it as data.

## Where real data lives instead

A real workbook lives outside the repo (e.g. `Documents\crisis-plan\<name>\`), and the
generated plan files land beside it. Google config lives in `%APPDATA%\crisis-plan\`
(or wherever `CRISISPLAN_GOOGLE_CONFIG` points). The tool **refuses** to read a real
workbook from, or write plan output into, the repo working tree — only `samples/`
(dummy) and `private/` (gitignored scratch) are allowed inside.

## Why the repo isn't a safe place for data

- **Git history is permanent.** A sensitive file committed and later "deleted" still
  lives in history forever. The only safe rule is: never commit it once.
- **This repo has a public remote.** Anything committed here must be safe for the
  whole internet to read, forever.
- **AI processing.** Any file an AI assistant reads is sent to its servers. Keeping
  real data out of the repo means an assistant working on this code never sees it —
  by design, not by trusting a policy.

## `.gitignore`

The `.gitignore` blocks spreadsheets, data exports, key/credential files, and a
`/private/` folder as a safety net. It is a backstop, **not** a licence to keep real
data in the folder — keep it out entirely.

## If sensitive data was ever committed

Stop and fix the history before pushing anywhere — deleting the file in a new commit
is **not** enough. Use `git filter-repo` / BFG, or recreate the repo fresh.
