# Crisis-plan generator (`tools/`)

This tool turns a person's data spreadsheet + a trigger (death) date into a
dated crisis plan: a **checklist** of what to do, in what order, and a
**calendar file** with the same tasks. It exists so nobody has to work that
out under grief — you run one command and follow the list.

**Written for a non-author.** If you've never seen this repo before, this
page is enough. (Developers: [`SPEC.md`](SPEC.md) is the specification,
[`PLAN.md`](PLAN.md) the build plan, [`../PRIVACY.md`](../PRIVACY.md) the
data rules.)

## I just need the plan, right now

You may not need to run anything:

1. **A printed copy of the latest plan is kept with the will** — check there
   first (the `key_documents` tab of the spreadsheet says where the will is).
2. On this computer, the current plan lives **in the same folder as the
   spreadsheet** — one folder per person, e.g.
   `Documents\crisis-plan\<name>\checklist.md`. Open it
   (any text editor works) and work top to bottom; `run-info.txt` says
   when it was generated. (Sample/demo runs go to `private\runs\<date>\`
   inside this repo instead.)

## Running it

### Quick demo (no setup decisions, safe)

Double-click **`demo.bat`** in the folder above this one. It runs entirely
offline against the built-in **dummy** sample data with a dummy date, and
writes its output to `private\runs\<today>\`. Use this to check the tool
still works.

### A real run

The real spreadsheet (e.g. `your-name.xlsx`) is **not in this folder** — it
lives outside the repo (see [`../PRIVACY.md`](../PRIVACY.md); ask where, or
check the safe / password manager). From a terminal in the repo folder:

```
python tools\generate.py -f C:\path\to\your-name.xlsx --date YYYY-MM-DD
```

where `--date` is the date of death. That's all that's required. The plan
files land in the same folder as the spreadsheet, regenerated on every
run (one workbook per folder; wizard edit backups go to `backups\`).

Options: `--out DIR` to write somewhere else
(must be outside this repo, or under its `private\`); `--mode {local,test,real}` for the Google
calendar push (`local` = offline, the default; `test` pushes into your
private test calendar, `real` into the shared family one — both need the
one-time setup in [`GOOGLE-SETUP.md`](GOOGLE-SETUP.md)); `--fresh` uses a
disposable, loudly-named test calendar instead. A Google failure is only
ever a WARN — the local files above are always written first. The tool
**refuses to run** if you point `-f` or `--out` at a tracked part of the
repo — that's deliberate, to keep real data out of git.

### What you get

Each run writes a dated snapshot folder containing:

- **`checklist.md`** — the human-readable plan, grouped by date, with day
  offsets from the trigger date, checkboxes, and who owns each task. Print
  it; keep a copy with the will.
- **`checklist-<name>.md`** — one personal checklist per person, when the
  work is shared between more than one (driven by the `crisis_roles`
  column on the household tab).
- **`plan.ics`** — the same tasks as a calendar. Import into Google
  Calendar / Outlook / Apple Calendar (File → Import).
- **`funeral-details.md` + `funeral-details.ics`** — a separate
  share-with-friends pair, generated only once the funeral tab has a
  service date (fill it in when known and re-run). Built from the funeral
  arrangements alone — nothing else from the spreadsheet can appear in
  it — so it's safe to forward to the wider circle.
- **`preparation.md`** — the *other* half: what to do **before** anything
  happens. Your `preparation` tab's actions/preferences (status tracked
  in the spreadsheet) plus gaps the tool notices — no will or LPA
  recorded, pensions without an expression of wish, too many accounts.
  Local only; never on any calendar.
- **`run-info.txt`** — what ran, from which file, with which date.
- **`google-preview.md`** — a plain-English list of exactly what a
  Google Calendar push would publish (titles are anonymised per the
  spreadsheet's `online_visibility` column; withheld rows are listed by
  number only). Written on **every** run, even fully offline ones —
  read it before ever turning the Google push on.

Every run ends with a **health summary** (per-stage OK/FAIL). If a stage
says FAIL, something has rotted — the message says what.

### Optional: look up bereavement links (`augment.py`)

For rows you've marked **`ai_visible?` = Y** in the spreadsheet, an
optional extra step can add practical detail lines (bereavement phone
lines, URLs) to the checklist:

```
python tools\augment.py -f C:\path\to\your-name.xlsx --provider manual
```

It first **shows you the exact queries** — only the institution's name
plus a fixed phrase, never owners, account labels, or locations — and does
nothing until you type `yes`. With `--provider manual` (the recommended
flow) *you* then search the web in your own browser and paste the useful
lines back in; nothing goes online from the tool at all. Results are
saved next to your spreadsheet (`your-name.augment.json`) and merged into the
checklist on the next plan run, marked `[ai]`. The calendar file never
carries them, the spreadsheet is never modified, and deleting the
`.augment.json` file cleanly removes the extras. `--list` previews the
queries without doing anything.

### Google setup (not required — for the future calendar push)

The Google Calendar push (`--mode test`/`real`) is not built yet. When it
is, its settings live in one small file **outside** this repo:
`%APPDATA%\crisis-plan\google.json` (or wherever the
`CRISISPLAN_GOOGLE_CONFIG` environment variable points). Copy
`tools\google.config.example.json` there and fill it in during the
one-time setup — calendar IDs and credential file paths only, and the
credential files themselves also live outside the repo. The health
summary's `google config` line tells you whether it was found. Until
then, every run still writes `google-preview.md` so you can see exactly
what a push *would* publish.

### Another household

The engine is data-driven — one household = one workbook. The friendly
way to start one is the **wizard** — a guided form in your browser
(entirely on this computer; nothing goes online):

```
python tools\wizard.py
```

Answer what you can (every section is optional), tell it where to save,
and it writes a ready-to-use workbook — then shows the exact command to
generate that household's plan. It can optionally run the one-time
Google Calendar setup at the end. It never overwrites an existing file.

Prefer a blank spreadsheet to fill in by hand?

```
python tools\make_template.py --blank --out C:\somewhere-outside-the-repo\alex.xlsx
```

That writes a clean template (instructions tab included, no dummy data).
Fill it in — the instructions tab explains every column — then run it
exactly like any other workbook:

```
python tools\generate.py -f C:\...\alex.xlsx --date YYYY-MM-DD
```

For the Google push, each household gets its **own calendars and its own
config file**: run `gsetup.py` with `CRISISPLAN_GOOGLE_CONFIG` pointing at
that household's config (e.g. `%APPDATA%\crisis-plan\google-alex.json`),
then pass `--google-config` (or the same env var) on their runs. Shared
task-status across households is deliberately **not** wired up — that's
a family consent conversation first (see PLAN.md 4.1).

## Safe to re-run, always

- The tool **never writes to the spreadsheet**. Anything typed in it is safe.
- Running twice produces the same plan — **no duplicates**. Each task has a
  stable ID, so re-importing `plan.ics` **updates** existing calendar
  entries rather than duplicating them.
- Got the date wrong? Re-run with the corrected `--date`: the same tasks
  simply move to the right days.
- The spreadsheet is the source of truth: fix a detail there, re-run, and
  the affected tasks update in place.

## If it won't run (setup from nothing)

1. Install **Python 3** (built/tested on 3.12) from python.org — tick
   "Add python to PATH" in the installer.
2. In a terminal: `pip install -r tools\requirements.txt`
   (one pinned dependency, `openpyxl`, for reading `.xlsx`; calendars and
   the Google client are standard library only). **No internet?** The
   wheels are committed in this repo:
   `pip install --no-index --find-links tools\vendor -r tools\requirements.txt`
3. Double-click `demo.bat`. If the demo works, the tool works.

Worst case — no Python, no working computer — the plan **does not depend on
this tool**: use the latest printed checklist, or open the newest
`private\runs\` folder from any backup.

## Ground rules (privacy — read before developing)

- Build and test **only** against `samples/pat.sample.xlsx` (dummy data).
  Never open or request a real workbook.
- The tool **reads** the `.xlsx`; it never writes back to it.
- No real data, secrets, or calendar IDs in any tracked file — see
  [`../PRIVACY.md`](../PRIVACY.md).

## Layout

- `tools/` — code (this directory).
- `samples/` — dummy spreadsheets, safe to commit and for AI to read.
- `private/` — generated output (`runs/<date>/`); **gitignored**, never tracked.
