"""Regression suite from build step 3.3: reconcile engine, test mode
(AC-I1, I2, I3, I5, I6, AC-M1, M2, M4).

Chains the real CLI: gsetup (fake transport) -> generate.py --mode test.
The fake stores state + a request log on disk, so every Google-side
claim is asserted against what the reconcile engine actually sent.
"""

import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SAMPLE = REPO / "samples" / "pat.sample.xlsx"
SCRATCH = Path(tempfile.mkdtemp(prefix="crisisplan-verify33-"))
CONFIG = SCRATCH / "google.json"
STATE = SCRATCH / "gcal-state.json"
ENV = dict(os.environ, CRISISPLAN_GOOGLE_CONFIG=str(CONFIG),
           CRISISPLAN_GCAL_FAKE=str(STATE))

failures = []


def check(name, ok, detail=""):
    print(f"  [{'PASS' if ok else 'FAIL'}] {name}" + (f" — {detail}" if detail else ""))
    if not ok:
        failures.append(name)


def run(script, *args):
    r = subprocess.run(
        [sys.executable, str(REPO / "tools" / script), *map(str, args)],
        capture_output=True, text=True, encoding="utf-8",
        errors="replace", env=ENV)
    return r


def run_cli(src, trigger, out, *extra):
    r = run("generate.py", "-f", src, "--date", trigger, "--out", out,
            "--mode", "test", *extra)
    if r.returncode != 0:
        print(r.stdout, r.stderr)
        raise SystemExit("CLI failed")
    return r.stdout


def state():
    return json.loads(STATE.read_text(encoding="utf-8"))


def test_events():
    s = state()
    cfg = json.loads(CONFIG.read_text(encoding="utf-8"))
    evs = s["events"].get(cfg["test_calendar_id"], {})
    return {e["extendedProperties"]["private"]["planKey"]: e
            for e in evs.values()
            if e.get("extendedProperties", {}).get("private", {})
            .get("planKey")}


run("gsetup.py", "--emails", "", "--household", "Sample family")

print("AC-I1: first push creates, second skips — no duplicates")
out = run_cli(SAMPLE, "2030-01-15", SCRATCH / "run-1")
# 71 plan events + the meta/annual-review self-canary (step 3.5) = 72.
check("all events created on first push",
      "72 created, 0 updated, 0 skipped" in out, out.splitlines()[-3])
evs1 = test_events()
check("72 tagged events in the test calendar", len(evs1) == 72)
out = run_cli(SAMPLE, "2030-01-15", SCRATCH / "run-2")
check("second identical push skips everything",
      "0 created, 0 updated, 72 skipped, 0 orphan-flagged" in out)
check("event count unchanged", len(test_events()) == 72)

print("AC-M2: no notifications ever, test calendar only")
log = state()["log"]
event_writes = [r for r in log if "/events" in r["path"]
                and r["method"] in ("POST", "PATCH")]
check("every event insert/patch sent sendUpdates=none",
      event_writes
      and all(r["params"].get("sendUpdates") == "none"
              for r in event_writes), f"{len(event_writes)} writes")
cfg = json.loads(CONFIG.read_text(encoding="utf-8"))
check("real calendar untouched by test mode",
      not state()["events"].get(cfg["real_calendar_id"]))

print("AC-I2: field edit updates in place")
from openpyxl import load_workbook  # noqa: E402

copy = SCRATCH / "edited.xlsx"
shutil.copyfile(SAMPLE, copy)
wb = load_workbook(copy)
ws = wb["banks"]
header = {c.value: i + 1 for i, c in enumerate(ws[1]) if c.value}
ws.cell(row=2, column=header["location_of_details"]).value = \
    "MOVED: green folder (reconcile test)"
wb.save(copy)
out = run_cli(copy, "2030-01-15", SCRATCH / "run-3")
check("exactly the edited row's event updated",
      "0 created, 1 updated" in out, out.splitlines()[-3])
evs = test_events()
check("same event id, new description",
      evs["banks/example-bank-plc-current-account/notify"]["id"]
      == evs1["banks/example-bank-plc-current-account/notify"]["id"]
      and "green folder" in
      evs["banks/example-bank-plc-current-account/notify"]["description"])

print("AC-I3: corrected date reschedules everything, same ids")
out = run_cli(SAMPLE, "2030-02-01", SCRATCH / "run-4")
check("all plan events updated, none created (canary untouched)",
      "0 created, 71 updated" in out and "1 skipped" in out)
evs = test_events()
check("ids stable across the reschedule",
      {k: e["id"] for k, e in evs.items()}
      == {k: e["id"] for k, e in evs1.items()})
check("dates actually moved",
      evs["base/register-death"]["start"]["date"] == "2030-02-06")

print("AC-I6: human status survives reconcile")
s = state()
test_cal = cfg["test_calendar_id"]
target = next(e for e in s["events"][test_cal].values()
              if e["extendedProperties"]["private"]["planKey"]
              == "base/register-death")
target["summary"] = "✓ " + target["summary"]
target["colorId"] = "5"
STATE.write_text(json.dumps(s), encoding="utf-8")
run_cli(SAMPLE, "2030-01-15", SCRATCH / "run-5")   # date change -> update
evs = test_events()
done = evs["base/register-death"]
check("done-mark ✓ preserved through an update",
      done["summary"].startswith("✓ ")
      and "Register the death" in done["summary"])
check("human colorId never touched", done.get("colorId") == "5")
check("plan fields still updated underneath",
      done["start"]["date"] == "2030-01-20")

print("AC-I5: orphans flagged, never deleted; untagged untouched")
s = state()
s["events"][test_cal]["manual-1"] = {
    "id": "manual-1", "summary": "Dentist (human's own event)"}
STATE.write_text(json.dumps(s), encoding="utf-8")
noprop = SCRATCH / "row-removed.xlsx"
shutil.copyfile(SAMPLE, noprop)
wb = load_workbook(noprop)
ws = wb["digital"]
header = {c.value: i + 1 for i, c in enumerate(ws[1]) if c.value}
for col_name, col in header.items():
    if col_name not in ("id", "online_visibility", "ai_visible?"):
        ws.cell(row=4, column=col).value = None   # drop the social row
wb.save(noprop)
out = run_cli(noprop, "2030-01-15", SCRATCH / "run-6")
check("removed row's event orphan-flagged",
      "1 orphan-flagged" in out, out.splitlines()[-3])
evs = test_events()
orphan = evs["digital/3/action"]
check("orphan kept with a visible ⚠ flag",
      orphan["summary"].startswith("⚠ [not in plan] ")
      and orphan["extendedProperties"]["private"]["orphaned"] == "true")
out = run_cli(noprop, "2030-01-15", SCRATCH / "run-7")
check("already-flagged orphan left alone on re-run",
      "0 orphan-flagged" in out and "52 skipped" not in out)
check("untagged human event untouched, unflagged",
      state()["events"][test_cal]["manual-1"]
      == {"id": "manual-1", "summary": "Dentist (human's own event)"})
out = run_cli(SAMPLE, "2030-01-15", SCRATCH / "run-8")
check("row back -> orphan unflagged and updated in place",
      "1 updated" in out
      and not test_events()["digital/3/action"]["summary"]
      .startswith("⚠")
      and "orphaned" not in
      test_events()["digital/3/action"]["extendedProperties"]["private"])

print("AC-M1: local mode never touches the transport")
before = STATE.read_bytes()
r = run("generate.py", "-f", SAMPLE, "--date", "2030-01-15",
        "--out", SCRATCH / "run-local", "--mode", "local")
check("local run OK with google SKIP",
      r.returncode == 0 and "local mode — no network" in r.stdout)
check("fake-transport state byte-identical after a local run",
      STATE.read_bytes() == before)

print()
if failures:
    raise SystemExit(f"FAILED: {failures}")
print("All step-3.3 checks passed (72 events reconciled).")
