# Build plan — crisis-plan generator

Durable, resumable build plan. **Designed so you can `/clear` between any two steps.**
All state that matters lives on disk: this file (step status), git history (checkpoints),
`tools/SPEC.md` (what to build), and the auto-memory index. A fresh session needs nothing
from the prior conversation.

---

## ▶ Resume protocol (read this first, every fresh session)

1. Read this file top to bottom, then [`SPEC.md`](SPEC.md) and [`../PRIVACY.md`](../PRIVACY.md).
2. Run `git log --oneline -15` to see what's already committed.
3. Find the **first `☐` step** in "Steps" below — that's the next unit of work.
4. Do exactly that step. Verify it against the acceptance criteria it names (from SPEC).
5. Tick it `☑`, update "Progress at a glance", and **commit** (see convention below).
6. Stop, or `/clear` and repeat.

**Suggested resume prompt after `/clear`:**
> "Continue the crisis-plan build per `tools/PLAN.md` — do the next unchecked step."

---

## Ground rules (hold on every step)

- **Privacy:** build and test **only against `samples/pat.sample.xlsx` (dummy data).**
  Never open or request a real workbook. No real data, secrets, or calendar IDs in
  any tracked file — see [`../PRIVACY.md`](../PRIVACY.md).
- **Read-only source:** the tool must never write to the input `.xlsx` (AC-P1).
- **Test what you build:** run it. Don't mark a step done on unrun code.
- **Regression suites:** `python tools/tests/run_all.py` after every step (sample-only,
  temp-dir outputs). A step that breaks an earlier suite isn't done. New steps with
  testable ACs add a `tools/tests/verify_<step>.py`.
- **One step = one commit.** The commit is the `/clear`-safe checkpoint.
- **The script reads the xlsx; humans own it.** Outputs regenerate; inputs never mutate.

## Conventions

- **Status:** `☐` todo · `▣` in progress · `☑` done · `⏸` blocked (note why).
- **Commit message:** `build: <step-id> — <short desc>` + the standard co-author trailer.
- **Paths:** code in `tools/`, dummy data in `samples/`, generated output in `private/`
  (gitignored). Calendar IDs/creds (Phase 3) live outside the repo.
- **Environment:** Python 3 + `openpyxl`. Pin versions in `tools/requirements.txt`.
- **Traceability:** each step names the SPEC acceptance criteria it satisfies.

---

## Progress at a glance

- Phase 0 — Setup: `☑` 0.1
- Phase 1 — Template + local engine: `☑` 1.1 `☑` 1.2 `☑` 1.3 `☑` 1.4 `☑` 1.5 `☑` 1.6 — **Phase 1 complete**
- Phase 2 — Remaining modules + visibility: `☑` 2.1 `☑` 2.2 `☑` 2.3 `☑` 2.4 `☑` 2.5 `☑` 2.6 — **Phase 2 complete**
- Phase 3 — Google (collaborative): `☑` 3.1 `☑` 3.2 `☑` 3.3 `☑` 3.4 `☑` 3.5 — **code complete**; *the real one-time Google run is still pending: `tools/GOOGLE-SETUP.md`*
- Phase 4 — Other households: `☑` 4.1 *(shared-status backend ⏸ awaiting family consent conversation)*
- Phase 5 — Onboarding wizard: `☑` 5.1 `☐` 5.2 *(pluggable backends — needs the family consent conversation first)* `☑` 5.3 *(pre-death prep plan, first slice)* `☐` 5.4 *(richer completeness rules)*

---

## Steps

### Phase 0 — Setup

**☑ 0.1 — Toolchain & skeleton**
- Verify Python 3 + `openpyxl` available; pick the `.ics` approach (stdlib string-build or
  a small vendored lib). Create `tools/requirements.txt` (pinned).
- Create `tools/` and `samples/`; add a `tools/README.md` stub. Confirm `private/` is
  gitignored (it is).
- **Verify:** `python --version` and an `import openpyxl` smoke test succeed.
- **Satisfies:** groundwork for D5.

### Phase 1 — Template workbook + local engine

**☑ 1.1 — Template workbook generator → `samples/pat.sample.xlsx`**
- A generator script builds the workbook: all tabs from SPEC §3, protected sheet with only
  input cells unlocked, dropdowns on constrained fields, an instructions tab, dummy rows.
- **Verify:** file opens; structure matches SPEC §3–4; `grep`-clean of real data.
- **Satisfies:** C-family (config surface), P2.

**☑ 1.2 — Tolerant reader**
- Parse the workbook into an internal model. Presence+content detection (tab counts only
  if ≥1 data row). Derive `planKey` per row (tab + slug). Ignore unknown tabs/columns.
- **Verify:** reads the sample; adding a junk tab/column doesn't error; never writes to the
  file.
- **Satisfies:** C1, C2, P1.

**☑ 1.3 — Task-module library (first slice)**
- Base/universal tasks + modules for **`banks`, `property`, `funeral_contacts`**.
  Implement `joint?`/`owner` branching and death-relative date math.
- **Verify:** correct tasks derived from the sample; toggling `joint?` changes output.
- **Satisfies:** C1, C3, AC-1 (partial).

**☑ 1.4 — Output generators (markdown + `.ics`)**
- Human-readable checklist (markdown) + `.ics` with `UID = planKey` (stable).
- **Verify:** `.ics` parses in a calendar app; dates correct relative to `--date`.
- **Satisfies:** D1 (partial).

**☑ 1.5 — CLI, archive, health-check, guards, `.bat`**
- `argparse`: `-f --date --mode --fresh --out`. Dated snapshot to `private/runs/<date>/`.
  Per-stage health summary. Guard: refuse if `-f`/`--out` inside the repo. `local` mode
  makes no network calls. `demo.bat` with safe defaults (sample, `local`, dummy date).
- **Verify:** double-click `demo.bat` runs end-to-end offline and writes a snapshot.
- **Satisfies:** P3, D1, D2, D3, D4, U1, U2, AC-1 (full).

**☑ 1.6 — Idempotency + cold-start README (Phase 1 done)**
- Run twice → outputs logically identical, no duplicates. Change a field → item updates by
  `UID`. Corrected date → reschedules. Write `tools/README.md` for a non-author.
- **Verify:** the above, by running.
- **Satisfies:** I1, I2 (local), I3, I4, D6.

### Phase 2 — Remaining modules

**☑ 2.1 — Flesh out the rest**
- Add task templates for `building_societies`, `cash_isas`, `shares_isas`,
  `share_portfolios`, `pensions`, `other_assets`, `wishes`, `key_documents`, `digital`,
  `access_pointers` (pointers only — reject secrets). Commit per module or small batch;
  update this file's checklist as you go.
- **Verify:** each module emits sensible tasks from sample rows.

**☑ 2.2 — Visibility layers (`id` / `online_visibility` / `ai_visible?`)** *(SPEC §13)*
- Template: add the three columns to every data tab (dropdowns, pre-filled ids,
  instructions text); regenerate the sample with mixed dummy values covering all three
  visibility levels. Reader: recognise the columns; absent columns behave as the
  defaults (`generic`, `N`). Tasks/outputs: each task carries a full and a redacted
  label; the `.ics` obeys the policy — redacted labels, `<tab>/<id>/<task-id>` UIDs,
  `none` rows suppressed; the checklist stays full-detail and gains the id ↔ label
  legend. Note: redacted rows change `.ics` UID identity — acceptable now, before the
  Phase-3 reconcile exists.
- **Verify:** sample rows at each level render correctly in both outputs; no real name
  appears anywhere in a `generic` row's event (label or UID); idempotency suite still
  passes.
- **Satisfies:** AC-V1, AC-V2, AC-V3; groundwork for AC-P4.

**☑ 2.3 — AI augmentation tool (`tools/augment.py`)** *(SPEC §13)*
- Minimised query extraction from `ai_visible? = Y` rows; print-and-confirm consent
  gate before any network call; pluggable lookup (offline stub + documented manual
  paste flow first — live search optional later); sidecar cache keyed by planKey in an
  untracked location; renderers merge cache lines when present.
- **Verify:** with the sample: the gate prints only institution/service names; cache
  round-trips into outputs; outputs are unchanged when the cache is absent; the
  workbook is never written.
- **Satisfies:** AC-V4, AC-V5.

**☑ 2.4 — Roles & personalised checklists** *(SPEC §14)*
- `household` gains multi-value `crisis_roles` (template + sample rows exercising held
  and unheld roles). Task templates get a default role (most `admin`; invite → `comms`,
  wishes → `organiser`/`medical`, etc.). Resolver: first holder, else executor → spouse
  → first person, with a visible fallback note per unheld role. Master checklist
  annotates assignees; per-person `checklist-<name>.md` when more than one assignee.
  `.ics` untouched.
- **Verify:** sample assigns every task; the deliberately-unheld role produces the
  fallback note; union of per-person checklists equals the master; editing roles
  changes no UIDs; idempotency + module + visibility suites still pass.
- **Satisfies:** AC-R1, AC-R2, AC-R3.

**☑ 2.5 — `funeral`, `professionals`, `life_insurance` modules + chronological tab order**
- Three new tabs + builders. A populated `funeral` tab *overrides* the generic
  funeral-director base task (contact the pre-arranged provider instead).
  `professionals` is type-driven (solicitor → notify + probate engagement; accountant →
  final tax). `life_insurance` covers policies and death-in-service. Reorder template
  tabs crisis-chronologically (reader is order-agnostic; presentational only).
- **Verify:** sample rows emit sensible tasks; the base-task override works; existing
  suites pass.
- **Satisfies:** AC-C1 extension to the new tabs.

**☑ 2.6 — Friends-facing funeral artifact (rerun flow)**
- A separate share-with-friends artifact (`funeral-details.md` + single-event `.ics`)
  generated **only** from the `funeral` tab, once real details are known via a re-run.
  Distinct audience = distinct artifact; no other tab's rows can leak into it.
- **Verify:** artifact contains only funeral-tab fields; absent/empty funeral tab →
  no artifact; reruns idempotent.
- **Satisfies:** SPEC §7 friends-artifact bullet.

### Phase 3 — Google integration (collaborative — needs your account)

**☑ 3.1 — Config scaffolding + publish preview** — loader for calendar IDs/creds from
outside the repo (env-var override; `%APPDATA%\crisis-plan\google.json` default; refuses
repo-tree paths); committed example config with dummy values; `google-preview.md`
written every run **before** any Google stage — the human-readable manifest of exactly
what `test`/`real` would publish (deliberately built ahead of the 3.2 plumbing).
Satisfies M6, M7, P2.
**☑ 3.2 — One-time setup command** — create + share the real calendar; establish the
persistent test calendar. Interactive with you. Satisfies M5.
*(Built + fake-transport tested. The actual one-time run against Google is pending —
follow `tools/GOOGLE-SETUP.md` when ready; stdlib-only OAuth, no new dependencies.)*
**☑ 3.3 — Reconcile engine (`test` mode)** — upsert by `extendedProperties.private.planKey`;
create/update/skip/flag-orphan; manage only tagged items; status read-only; notifications
off. Satisfies I1, I2, I5, I6, M1, M2, M4.
*(Done-mark convention: a leading `✓` on an event title survives updates; colours never
touched — see GOOGLE-SETUP.md.)*
**☑ 3.4 — `real` mode + `--fresh` + graceful degradation** — Google failure never blocks
local artifacts. Satisfies M3, M4, D3.
**☑ 3.5 — Durability hardening** — vendor deps, self-scheduling annual-review event,
complete the health-check. Satisfies D5.

### Phase 4 — Other households

**☑ 4.1 — Generalise** — reuse the engine with other households'
files; decide the shared-status backend (a consent conversation). Later.
*(Engineering done: `make_template.py --blank --out` for new households,
`--google-config` / one config + calendars per household, README walkthrough, second-
household suite incl. the sole-name-house asymmetry. The shared-status backend
decision is **⏸ deferred** — a family consent conversation, not a technical step.)*

### Phase 5 — Onboarding wizard (later, sketch in SPEC §15)

**☑ 5.1 — Form-style wizard** — local-only guided interview that writes a filled
workbook (fresh file outside the repo; never edits an existing one) and finishes by
running the one-time calendar setup + family shares. Questions driven by the template's
column/dropdown definitions so wizard and template can't drift. Design properly before
building; primarily for Phase-4 households.
*(`tools/wizard.py`: stdlib http.server on 127.0.0.1, form generated from
make_template TABS, repo-path guards, per-row consent checkboxes + live label
preview, safe visibility defaults, optional gsetup hand-off, first plan generated
beside the workbook with clickable links. Also an **editor**: load an existing
workbook, rows keep their ids (deleted ids never re-issued), every save writes a
timestamped backup first.)*
**☑ 5.3 — Pre-death preparation plan (first slice)** *(SPEC §16)* — deliberately
simple, built to demo to the family: a `preparation` tab (item / type
action·preference·note / owner / priority / status — status human-owned) + noticed
gaps derived from the rest of the workbook (no will/LPA, empty wishes, missing
pension/life nominations, no funeral plan, consolidation nudge) → local-only
`preparation.md` beside the workbook. No calendar, no dates math. Richer derivation
and a calendar layer stay future work.
**☐ 5.4 — Richer plan-completeness analysis** — grow the deterministic gap rules
beyond the first seven: insurance nominated-to-partner nuances (partner predeceasing,
joint-life policies), stale-review nudges (workbook untouched for a year), and
age/circumstance-aware guidance (e.g. account consolidation — a judgment that needs
age or context data the workbook doesn't hold yet, plus a decision on whether it
should). Keep every rule deterministic and self-clearing; still no AI at runtime.
**☐ 5.2 — Pluggable data backend (app)** — per-household choice: local encrypted
workbook (default, keeps the privacy posture) OR an opt-in shared Google Sheet the
wizard fills (a consent decision — puts data in cloud custody; see SPEC §15).
Backend = where the workbook lives; Sheet backend exports to `.xlsx` into the same
tolerant reader, engine unchanged.

---

## After each step (the `/clear`-safe ritual)

1. Verify against the step's named ACs by **running** the tool.
2. Flip the step and the "Progress at a glance" line to `☑`.
3. `git add -A && git commit -m "build: <step-id> — <desc>"` (+ co-author trailer) and push.
4. Now it's safe to `/clear`.
