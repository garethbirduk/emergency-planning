"""Regression suite from build step 1.6: idempotency ACs (I1–I4).

Runs tools/generate.py end-to-end (subprocess, the real CLI) against the
DUMMY sample only, writing outputs to a temp dir outside the repo, and
checks:
  AC-I1  two identical runs -> byte-identical checklist.md + plan.ics,
         no duplicate UIDs
  AC-I2  changing a non-key plan field (on a temp COPY of the sample)
         -> same UID set, that event's content updated in place
  AC-I3  corrected trigger date -> same UID set, every DTSTART shifted
  AC-I4  the input .xlsx hash is unchanged after all runs
"""

import hashlib
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SAMPLE = REPO / "samples" / "pat.sample.xlsx"
SCRATCH = Path(tempfile.mkdtemp(prefix="crisisplan-verify16-"))
CLI = REPO / "tools" / "generate.py"

DATE_A = "2030-01-15"
DATE_B = "2030-02-01"          # 17 days later
SHIFT_DAYS = (date.fromisoformat(DATE_B) - date.fromisoformat(DATE_A)).days

failures = []


def check(name, ok, detail=""):
    print(f"  [{'PASS' if ok else 'FAIL'}] {name}" + (f" — {detail}" if detail else ""))
    if not ok:
        failures.append(name)


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def run_cli(src, trigger, out_dir):
    r = subprocess.run(
        [sys.executable, str(CLI), "-f", str(src), "--date", trigger,
         "--out", str(out_dir)],
        capture_output=True, text=True)
    if r.returncode != 0:
        print(r.stdout)
        print(r.stderr)
        raise SystemExit(f"CLI failed for {out_dir}")
    return r.stdout


def parse_ics(path):
    """Unfold and parse plan.ics -> {uid: {prop: value}} per VEVENT."""
    text = Path(path).read_bytes().decode("utf-8")
    unfolded = text.replace("\r\n ", "")
    events, cur = {}, None
    for line in unfolded.split("\r\n"):
        if line == "BEGIN:VEVENT":
            cur = {}
        elif line == "END:VEVENT":
            events[cur["UID"]] = cur
            cur = None
        elif cur is not None and ":" in line:
            prop, value = line.split(":", 1)
            cur[prop.split(";")[0]] = value
    return events


sample_hash_before = sha256(SAMPLE)

# --- AC-I1: run twice, identical inputs --------------------------------------
print("AC-I1: two identical runs")
run_a, run_b = SCRATCH / "run-a", SCRATCH / "run-b"
run_cli(SAMPLE, DATE_A, run_a)
run_cli(SAMPLE, DATE_A, run_b)
check("checklist.md byte-identical",
      (run_a / "checklist.md").read_bytes() == (run_b / "checklist.md").read_bytes())
check("plan.ics byte-identical",
      (run_a / "plan.ics").read_bytes() == (run_b / "plan.ics").read_bytes())

events_a = parse_ics(run_a / "plan.ics")
uid_lines = re.findall(rb"^UID:", (run_a / "plan.ics").read_bytes(), re.M)
check("no duplicate UIDs", len(uid_lines) == len(events_a),
      f"{len(uid_lines)} events, {len(events_a)} unique UIDs")
check("plan is non-trivial", len(events_a) >= 10, f"{len(events_a)} events")

# --- AC-I3: corrected trigger date reschedules -------------------------------
print("AC-I3: corrected trigger date")
run_c = SCRATCH / "run-c"
run_cli(SAMPLE, DATE_B, run_c)
events_c = parse_ics(run_c / "plan.ics")
check("same UID set (no dupes, no strays)",
      set(events_a) == set(events_c))
shifted = all(
    (date(int(d[:4]), int(d[4:6]), int(d[6:8])) -
     date(int(a[:4]), int(a[4:6]), int(a[6:8]))).days == SHIFT_DAYS
    for uid in events_a
    for a, d in [(events_a[uid]["DTSTART"], events_c[uid]["DTSTART"])])
check(f"every DTSTART shifted by exactly {SHIFT_DAYS} days", shifted)

# --- AC-I2: change a non-key plan field on a scratch copy --------------------
print("AC-I2: plan-field change updates in place")
from openpyxl import load_workbook  # noqa: E402

copy = SCRATCH / "pat.modified.xlsx"
shutil.copyfile(SAMPLE, copy)
wb = load_workbook(copy)
ws = wb["banks"]
header = {c.value: i + 1 for i, c in enumerate(ws[1]) if c.value}
col = header["location_of_details"]
target_row = next(r for r in range(2, ws.max_row + 1)
                  if ws.cell(row=r, column=header["bank"]).value)
old_val = ws.cell(row=target_row, column=col).value
NEW_VAL = "MOVED: blue box file, spare-room shelf (test edit)"
ws.cell(row=target_row, column=col).value = NEW_VAL
wb.save(copy)
print(f"  (edited banks row {target_row}: location_of_details "
      f"{old_val!r} -> {NEW_VAL!r})")

run_d = SCRATCH / "run-d"
run_cli(copy, DATE_A, run_d)
events_d = parse_ics(run_d / "plan.ics")
check("UID set unchanged by the field edit", set(events_a) == set(events_d))

changed = [uid for uid in events_a
           if events_a[uid] != events_d[uid]]
notify_uid = next((u for u in changed if "/notify" in u), None)
check("edited row's event(s) updated under the SAME UID",
      notify_uid is not None and NEW_VAL.replace(",", "\\,")
      in events_d[notify_uid]["DESCRIPTION"],
      f"changed events: {changed}")
check("only the edited row's events changed",
      all(uid.startswith("banks/") for uid in changed),
      f"{len(changed)} changed, all under the edited banks row")

# --- AC-I4 (checklist): human ticks survive regeneration ---------------------
print("AC-I4: checklist ticks survive re-runs")
md_path = run_a / "checklist.md"
text = md_path.read_text(encoding="utf-8")
target = next(ln for ln in text.splitlines()
              if ln.startswith("- [ ] ") and "Tell Us Once" in ln)
md_path.write_text(text.replace(target,
                                target.replace("- [ ]", "- [x]"), 1),
                   encoding="utf-8")


def ticked_tell_us_once(path):
    return any(ln.startswith("- [x] ") and "Tell Us Once" in ln
               for ln in path.read_text(encoding="utf-8").splitlines())


run_cli(SAMPLE, DATE_A, run_a)
regen = md_path.read_text(encoding="utf-8")
check("ticked task still ticked after regeneration",
      ticked_tell_us_once(md_path))
check("only that task is ticked", regen.count("- [x] ") == 1)
check("tick propagated to the assignee's personal checklist",
      ticked_tell_us_once(run_a / "checklist-alex-sample.md"))
run_cli(SAMPLE, DATE_B, run_a)   # even a rescheduled plan keeps it
check("tick survives a changed trigger date",
      ticked_tell_us_once(md_path))

# --- AC-I4: input workbook never modified ------------------------------------
print("AC-I4: input untouched")
check("sample .xlsx hash unchanged after all runs",
      sha256(SAMPLE) == sample_hash_before)

print()
if failures:
    raise SystemExit(f"FAILED: {failures}")
print(f"All step-1.6 checks passed ({len(events_a)} events per plan).")
