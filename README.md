# Emergency planning — crisis-plan generator

A local, durable, re-runnable tool that turns one person's data spreadsheet plus a
trigger (death) date into a **dated crisis plan**: a task checklist of what to do and
in what order, and a calendar file with the same tasks. It exists so nobody has to
work that out under grief — you run one command and follow the list.

- **One engine, data-driven.** One household = one workbook. The same engine serves
  any number of people; nothing about a household is hard-coded.
- **Local and private by design.** The tool runs offline; the real workbook lives
  *outside* this repo and never touches git. See [`PRIVACY.md`](PRIVACY.md).
- **Durable.** Pinned, vendored dependencies (openpyxl only), stdlib `.ics` output,
  no framework — built to still run on a bare Python install a decade from now.

## Try it (safe, offline, dummy data)

Double-click **`demo.bat`** (Windows), or:

```
python tools/generate.py -f samples/pat.sample.xlsx --date 2030-01-15 --mode local
```

Output lands in `private/runs/<today>/` (gitignored): `checklist.md`, per-person
checklists, `plan.ics`, `preparation.md`, and a preview of what a calendar push
would publish.

## Start a real plan

The friendly way is the **wizard** — a guided form in your browser, served from a
local process on your own machine (nothing goes online):

```
python tools/wizard.py
```

Prefer a spreadsheet? `python tools/make_template.py --blank --out <path>` writes a
clean template with an instructions tab. Either way, keep the real workbook
**outside** this repo — see [`PRIVACY.md`](PRIVACY.md).

Full usage, including the optional Google Calendar push and the AI-free augment
flow: [`tools/README.md`](tools/README.md). The specification is
[`tools/SPEC.md`](tools/SPEC.md); the build history is [`tools/PLAN.md`](tools/PLAN.md).

## Tests

```
python tools/tests/run_all.py
```

Sixteen regression suites, run against the committed dummy sample only; they write
to temp directories and never touch a real workbook.

## Requirements

Python 3.12+ and openpyxl. Offline install from the vendored wheels:

```
pip install --no-index --find-links tools/vendor -r tools/requirements.txt
```
