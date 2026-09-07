# Crisis-plan generator — specification

A local, durable, re-runnable tool that turns a person's data spreadsheet + a trigger
(death) date into a dated crisis plan: a task checklist and a calendar. One engine,
data-driven per person/household — one workbook per person, any number of households.
First target: **the author's own household**, as both a dry run and a genuine safeguard.

Design was worked out in conversation on 2026-08-02; this is the agreed spec that the
build follows.

## 1. Purpose

- Remove cognitive load at the worst possible moment: the dated "what to do, in what
  order, who owns it" is pre-computed, not figured out under grief.
- Work for **any** household by swapping the data file — same engine.
- Still be runnable, or at least *readable*, **10+ years** from now.

## 2. Architecture principles

1. **Repo holds the plan/code, never data or secrets.** (See [`../PRIVACY.md`](../PRIVACY.md).)
2. **The `.xlsx` is the human-owned source of truth** (facts + state). The script
   **reads only; never writes back** → human-entered detail can never be lost on rerun.
3. **Outputs are regenerated idempotently** via stable keys.
4. **Layered by durability:** local plain-text outputs are produced first and
   unconditionally; Google is a best-effort add-on that can fail without harm.
5. **Data-driven config:** a module's tasks generate only if its tab exists *and* has
   data. Absent/empty tab = that asset class doesn't apply.
6. **Tolerant reader:** unknown tabs/columns are ignored, not errors — detail can be
   added freely without breaking the tool.
7. **Non-identifying externally:** anything sent to an external surface carries generic
   labels by default; real specifics stay local. A per-row visibility policy (§13) lets
   the human opt individual rows up (full label) or out (local-only).

## 3. Modules (spreadsheet tabs)

Each tab = a module = *input columns* + *task templates it emits* + *branching*.

- **Control:** `household` — members **and the wider circle** (the doers are often
  adult children + their spouses, not co-residents): who's invited, legal role
  (executor…), **functional crisis roles** for task ownership (§14), region.
- **Assets:** `banks` · `building_societies` · `cash_isas` · `shares_isas` ·
  `share_portfolios` · `pensions` · `property` · `other_assets` (cars, valuables) ·
  `life_insurance` (incl. death-in-service — the classic "nobody knew it existed"
  asset).
- **Non-asset:** `funeral_contacts` (who to invite) · `funeral` (provider, package,
  venue — pre-arms and *overrides* the generic funeral-director base task) ·
  `professionals` (solicitor, accountant, adviser… type-driven tasks) ·
  `access_pointers` (where to find logins — *pointers, never secrets*) · `wishes`
  (funeral/DNR/donation) · `key_documents` (will/deeds/LPA locations) · `digital`
  (subscriptions, domains).

Branching column present on asset tabs: **`joint?` / `owner`** — sole → probate path
(more tasks); joint → survivor path (fewer). This one column generalises the
sole-name house asymmetry the design was born from to any household.

## 4. Spreadsheet (`.xlsx`) design

- Real `.xlsx` (via `openpyxl`).
- **Protected sheet; only input cells unlocked** (accident-prevention, not security).
- **Dropdowns** (data validation) on constrained fields (`joint?`, type, Y/N).
- **Formulas only for the human's eyes** (totals, hints). The script derives all logic
  from raw input cells and never depends on Excel-computed values.
- **Instructions/README tab**; clear locked headers per tab.
- **Tab order is crisis-chronological** (people, wishes, funeral, documents, access,
  contacts, professionals, then assets, then digital) — purely presentational; the
  reader is order-agnostic.
- `access_pointers` holds **locations/hints only** — never passwords, never seed phrases.
- **Visibility columns** on every data tab (§13): `id` (short stable row reference,
  pre-filled by the template, human-owned), `online_visibility`
  (`full`/`generic`/`none`; blank = `generic`), `ai_visible?` (`Y`/`N`; blank = `N`).

## 5. Idempotency & reconcile

- **Stable `planKey` per row**, derived by the script (tab + slug) — the human never
  manages keys.
- Calendar items matched on **`extendedProperties.private.planKey`**, not the raw event
  id.
- **Reconcile rules:** create missing · update changed *plan* fields · skip unchanged ·
  **flag orphans (keep, never delete)**. Only items bearing our tag are managed;
  untagged items are left entirely alone.
- **Field ownership:** plan fields owned by the source (`.xlsx`); **status owned by the
  human/Google side** — the script reads status, never overwrites it.
- **Re-run with a corrected date reschedules** existing items (same keys); no duplicates.

## 6. Calendar modes

| Mode | Target | Notifications | Role |
|---|---|---|---|
| `local` | none (no Google) | — | Offline canary; `.ics` + plain files only |
| `test` | persistent test calendar (default) | **off** | Google-path canary; exercises the *real* upsert code |
| `real` | configured shared calendar | **off** on reconcile | Live run |

- **`--fresh`** creates a clearly-named disposable calendar (`⚠ TEST … DELETE ME`) for a
  clean-room test.
- **Invite model:** the real calendar is **shared once, at setup** (calendar membership,
  not per-event guests) → re-runs never spam anyone.
- **Setup is a separate, deliberate one-time step** (create + share the real calendar);
  normal runs never create or share it.
- **Calendar IDs and credentials** live in config outside the repo.

## 7. Outputs

- **Always, unconditionally, locally:** a plan snapshot containing a
  human-readable checklist (markdown/HTML) **and** an `.ics`, plus a personal
  `checklist-<name>.md` per assignee when tasks are shared out (§14). Default
  location: **the workbook's own folder** (one workbook per folder) for a real
  (outside-repo) workbook — the current plan, regenerated in place with stale
  files removed; wizard edit backups live in its `backups/` subfolder. The
  wizard makes that folder per person: the chosen root gains a subfolder named
  after the planholder (`…\crisis-plan\pat\`, `…\crisis-plan\alex\`), so
  several people's plans coexist under one root without colliding. Or
  dated `private/runs/<date>/` folders for in-repo sample/dev runs.
- **Friends-facing funeral artifact** (Phase 2.6): funeral details for the wider
  circle are a *separate* artifact generated **only** from the `funeral` tab — a
  distinct audience gets a distinct artifact; rows from other tabs can never leak
  into it. Regenerated on a re-run once real details (date, venue) are known.
- **Print step:** recommend a printed copy kept with the will (ultimate fallback).
- **Google push:** best-effort, after local outputs, never blocks them.
- **Google publish preview:** every run (all modes) writes
  `google-preview.md` — a human-readable manifest of **exactly** what the
  Google stage would create or update: the external titles, dates,
  descriptions and keys (the §13-redacted surface), written before any
  network stage exists or runs. Rows with `online_visibility = none`
  appear only as withheld references. The family can read this file to
  see precisely what would leave the house before the first `test`/`real`
  push.
- **Health summary** printed each run (per-stage success/failure).
- **The `.ics` counts as an external surface** (it exists to be imported into cloud
  calendars): the per-row visibility policy (§13) applies to it exactly as to Google.
  The markdown checklist is the always-full-detail local artifact.

## 8. Durability (10-year horizon)

- **Pinned dependencies** (`requirements.txt`); **vendored** critical libs for offline
  install; dependency-light core. Documented Python version.
- `.xlsx` is an ISO standard → durable. **Google/OAuth is the fragile part** (note:
  OAuth apps left in "Testing" expire refresh tokens in ~7 days — keep Google optional).
- **Self-health-check** each run tells you what has rotted while it's cheap to fix.
- **Annual re-run as canary**; the tool schedules its own "annual review" reminder.
- **README for cold-start** by a non-author (spouse/executor), with a pointer to the latest
  run and the printed copy.

## 9. Graceful degradation

- Local artifacts produced first and independently of Google.
- A fragile-tier (Google) failure is logged and skipped; the run still succeeds.
- Tolerant `.xlsx` reading (unknown tabs/columns ignored).

## 10. CLI & UX

```
generate.py  -f <person>.xlsx  --date <trigger>  --mode {local,test,real}  [--fresh]  [--out <dir>]
```

- **`.bat` wrappers** (`demo.bat`) for double-click; safe defaults (sample data, `local`
  mode, dummy date).
- **Guard:** refuse to run if `-f` or `--out` points inside the repo working tree.

## 11. Privacy & security

- Real `.xlsx` is gitignored, passed via `-f`; **AI/assistants only ever see the sample.**
- No secrets or calendar IDs in the repo; creds/IDs in the encrypted store.
- External (Google) labels are non-identifying.

## 12. Scope & phasing

- **Phase 1 (now):** the author's household, `local` mode, `.ics` + plain outputs + dated
  archive. Full tab template (unused tabs left empty); task templates fully wired for a
  first slice — **`banks`, `property`, `funeral_contacts`** — proving input→tasks→output.
- **Phase 2:** remaining modules.
- **Phase 3:** Google `test` then `real`; reconcile; the one-time setup step.
- **Later:** extend to further households (same engine, different `.xlsx`); decide the
  shared-status backend (a consent conversation, not a technical one).

## 13. Visibility layers & AI augmentation

Each row is an item in a human-readable database; two per-row columns give the human
item-level control over where its details may appear. Defaults are the safe choice, and
workbooks without these columns behave as the defaults (tolerant reader).

- **`id`** — short stable row reference (e.g. `3`). The template pre-fills `1, 2, 3…`;
  the human owns it thereafter and must never reuse a number after deleting a row. Used
  in redacted labels ("Building Society #3") and as the row's non-identifying external
  key. If a row needs redaction but has a blank `id`, the run emits a health warning and
  falls back to a short opaque hash of the slug (stable, but ugly — fill the id in).
- **`online_visibility`** — what external surfaces (`.ics`, Google) may show:
  - `full` → real label + id: "Inform Sampletown Building Society (#3)"
  - `generic` (default when blank) → generic tab label + id: "Inform Building Society #3"
  - `none` → the row's tasks appear in local outputs only; nothing external at all.
  For `generic`/`none` rows the external UID is `<tab>/<id>/<task-id>` — no name-derived
  slug may leak into external keys or extendedProperties.
- **`ai_visible?`** — `Y` = consent for the augmentation step to include this row;
  default `N`.

**Augmentation** (`tools/augment.py`, run by the household; entirely opt-in):

1. Collects one **minimised query** per consenting row — the institution/service name
   and task context only; never the owner, account label, or location fields.
2. **Prints the exact query list and stops for explicit confirmation** before any
   network call (the consent gate, visible to the whole family).
3. Looks up practical extras (bereavement URLs, phone lines) and caches them in a
   sidecar file outside git (`private/` or beside the real workbook), keyed by planKey.
4. Renderers merge cached augmentations as extra detail lines when the cache exists;
   everything works identically without it. The workbook is never written (P1 holds).

Dev-time AI (assistants building this tool) still only ever sees the sample workbook —
`ai_visible?` governs the household's runtime lookups, not development.

## 14. Roles & personalised checklists

The plan's promise is "what to do, in what order, **who owns it**" (§1). Ownership is
driven by *functional* crisis roles, not just legal ones:

- `household` gains a **`crisis_roles`** multi-value column (comma-separated).
  Vocabulary: `admin` (forms, banks, probate legwork) · `organiser` (funeral
  logistics) · `local` (in-person visits to close friends) · `medical`
  (donor/immediate medical decisions) · `comms` (contacting the wider circle).
  The legal `role` column stays (planholder/executor/spouse detection); `executor`
  also counts as a held role for assignment.
- **The vocabulary is an enum, held in one place** (`tools/roles.py`) and read from
  there by the engine, the template instructions and the wizard — which offers the
  roles as tick-boxes, so a misspelling can't be authored. A value the enum doesn't
  recognise is reported as a note in the plan, never treated as an unheld role:
  the failure it prevents is a typo silently routing someone's tasks to the
  fallback person while the plan still reads as correct. Only spelling variants of
  the same word (`organizer`, `administrator`, `communications`) are accepted
  quietly; nothing else is guessed at.
- **Every task template carries a default role** (most are `admin`).
- **Resolution:** the first listed person holding the role gets the task. A role held
  by nobody falls back to the executor, else the spouse, else the first listed person
  — **and the run says so in a visible note, never silently**. With nobody listed
  besides the planholder, tasks are simply unowned (the pre-§14 behaviour).
- **Outputs:** the master checklist annotates each task with its assignee; when more
  than one person has tasks, a personal `checklist-<name>.md` is written per
  assignee. The `.ics` is unchanged by assignment for now — assignee names in events
  would hand the family roster to an external surface; revisit with Phase 3.
- **Identity:** role/assignment edits change plan fields only; task UIDs/planKeys
  never change.

## 15. Onboarding wizard (later phase — sketch)

A form-style wizard that interviews a household and produces their filled
workbook + calendar setup, so nobody has to face a 17-tab spreadsheet cold.
Agreed direction, not yet designed in detail:

- **An authoring front-end, not a new store.** The wizard's output IS the
  workbook (outside the repo); the engine, spec and privacy model are
  unchanged. It can reopen a workbook it understands as an **editor**,
  but never destructively: every save first writes a timestamped backup
  beside the file, row ids survive edits and are never re-issued after a
  deletion (calendar identity depends on that), and hand-added unknown
  columns — which the form can't show — are the signal to edit in Excel
  instead. The generator still never writes the workbook (P1).
- **Local-only.** Answers are the sensitive data, so no cloud forms (a
  Google Form would hand the data to a third party before the plan even
  exists — see PRIVACY.md). Think stdlib-served local web page or guided
  CLI/TUI, offline.
- **Single source of truth for questions:** the template's column/dropdown
  definitions (make_template.py TABS) drive the wizard's questions, so the
  two can't drift.
- **Finishes with the calendar:** the last page wraps the one-time setup
  (§6 / step 3.2) — collect family emails, create + share the real
  calendar (membership shares, never per-event invites), write the config.
- **Audience:** other households first — it exists to
  lower their barrier to entry (Phase 4 reuse).

Further out, the wizard can grow into an app with a **pluggable data
backend**, chosen per household:

- **Local, optionally encrypted** (default; keeps today's privacy
  posture): the workbook at rest in an encrypted container / encrypted
  file, synced by any dumb file-sync if wanted — ciphertext is safe to
  sync; the tool never sees plaintext without the human unlocking it.
- **Shared Google Spreadsheet** (opt-in, eyes open): the wizard fills a
  Sheet the family co-edits. This puts the DATA in third-party cloud
  custody — exactly what PRIVACY.md currently rules out — so it is a
  per-household **consent decision** (same conversation as the Phase-4
  shared-status backend), never a default. Engine impact is small by
  design: export the Sheet to `.xlsx` (Drive API) and feed the same
  tolerant reader — the backend decides *where the workbook lives*, the
  engine stays identical.

## 16. Pre-death preparation tasks (later phase — sketch)

The workbook is anchored on one person — the household tab's
**planholder** row (the wizard enforces this). Today the engine only
answers "what do we do when they die"; a later phase adds the mirror:
**what the planholder should do now, while alive**, derived from the
same workbook:

- Gaps become tasks: no `key_documents` will row → "write a will"; empty
  `wishes` → "record funeral/DNR/donation wishes"; no LPA rows → "set up
  LPAs" (see `pre-work/` in this repo); blank `nomination_in_place?` →
  "file the pension expression of wish"; many single-institution rows →
  "consider consolidating accounts".
- These are **planholder-owned** tasks (unlike crisis tasks, which are
  everyone-else's), not death-relative — they'd be a separate checklist
  / calendar layer, likely alongside the annual-review canary.
- Same engine principles: read-only workbook, idempotent keys, local
  first.

*First slice built (step 5.3):* a `preparation` tab (actions /
preferences / notes with human-owned status) + a short list of derived
gap rules → local-only `preparation.md` beside the workbook. Calendar
layer and richer derivation remain future work.

---

## Acceptance criteria

### Data & privacy
- **AC-P1** A run never writes to, modifies, or deletes the input `.xlsx`.
- **AC-P2** No real personal data, secrets, or calendar IDs appear in any git-tracked
  file (grep-clean; `.gitignore` blocks spreadsheets/credentials).
- **AC-P3** The tool refuses to run, with a clear message, if `-f` or `--out` resolves to
  a path inside the repo working tree.
- **AC-P4** Anything sent to Google contains only non-identifying labels (no account
  numbers, balances, or account names).

### Source of truth / config
- **AC-C1** A module's tasks are generated **iff** its tab exists and has ≥1 populated row.
- **AC-C2** Unknown tabs and unknown columns are ignored without error.
- **AC-C3** The `joint?`/`owner` value changes which tasks are generated (sole → probate
  path; joint → survivor path), demonstrable from sample rows.

### Idempotency
- **AC-I1** Two runs with identical inputs produce logically identical outputs and **no
  duplicate** calendar entries.
- **AC-I2** Changing a plan field in the `.xlsx` and re-running **updates** the matching
  calendar item in place (by `planKey`), not a duplicate.
- **AC-I3** Re-running with a corrected trigger date **reschedules** existing items; no
  duplicates.
- **AC-I4** Any human-entered detail/status in the `.xlsx` is unchanged after any number
  of reruns.
- **AC-I5** A tagged calendar item absent from the plan is **flagged, never deleted**;
  untagged items are left untouched and unflagged.
- **AC-I6** Status fields are read, never overwritten.

### Outputs & durability
- **AC-D1** Every run writes a snapshot containing a human-readable checklist and
  an `.ics` — into the workbook's own folder for a real workbook, or to
  dated `private/runs/<date>/` for in-repo sample runs; `run-info.txt` records
  when it was generated.
- **AC-D2** In `local` mode the tool runs with **no credentials and no internet**.
- **AC-D3** Local artifacts are produced before, and independently of, any Google step; a
  Google failure does not prevent them.
- **AC-D4** Each run prints a per-stage health summary (success/failure).
- **AC-D5** Dependencies are pinned and an offline (vendored) install path is documented.
- **AC-D6** A README lets a non-author run the tool or locate the latest plan / printed
  copy.

### Calendar modes
- **AC-M1** `local` makes no Google API calls.
- **AC-M2** `test` writes only to a test calendar and sends **no notifications** to real
  people.
- **AC-M3** `real` writes to the configured shared calendar with notifications suppressed
  on reconcile.
- **AC-M4** Default `test` reconciles into a **persistent** test calendar (same upsert
  code as `real`); `--fresh` creates a clearly-named disposable calendar.
- **AC-M5** A normal run never creates or shares the real calendar; setup/sharing is a
  separate explicit step.
- **AC-M6** Calendar IDs and credentials are read from config **outside** the repo.
- **AC-M7** Every run, in every mode, writes `google-preview.md` listing exactly
  the events the Google stage would publish — external labels/keys only,
  matching the `.ics` surface one-to-one — before any Google call is possible;
  `online_visibility = none` rows appear only as withheld references.

### Visibility & augmentation
- **AC-V1** `online_visibility=full` → external labels carry the real name + id;
  `generic`/blank → a generic tab label + id with no real name anywhere in the event;
  `none` → the row is absent from external outputs entirely while still present in the
  local checklist.
- **AC-V2** For `generic`/`none` rows, external UIDs/keys contain no name-derived text —
  tab + id (+ task id) only.
- **AC-V3** The local checklist always carries full detail and includes an id ↔
  real-label legend covering every redacted row.
- **AC-V4** The augment tool reads only `ai_visible? = Y` rows, sends only
  institution/service names, and prints the exact queries requiring explicit
  confirmation before any network call.
- **AC-V5** Augmentation results live in an untracked sidecar cache; the workbook is
  never written; outputs render correctly with or without the cache.

### Roles & ownership
- **AC-R1** With at least one person besides the planholder listed, every task has an
  assignee; a role held by nobody produces a visible fallback note, never a silent
  default.
- **AC-R2** When more than one person has tasks, per-person checklists are written and
  the union of their tasks equals the master checklist's tasks.
- **AC-R3** Changing `crisis_roles` assignments changes assignees only — no
  UID/planKey changes.

### CLI / UX
- **AC-U1** `demo.bat` runs end-to-end on a double-click with safe defaults (sample
  data, `local` mode, dummy date).
- **AC-U2** The CLI supports `-f`, `--date`, `--mode`, `--fresh`, `--out`.

### Phase-1 completeness
- **AC-1** With the sample workbook, `generate.py` in `local` mode generates a correct dated
  plan covering **banks, property, and funeral_contacts** end to end.
