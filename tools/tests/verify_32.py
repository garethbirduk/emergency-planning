"""Regression suite from build step 3.2: one-time Google setup
(AC-M5, AC-M6).

Runs the REAL gsetup.py against the offline fake transport
(CRISISPLAN_GCAL_FAKE): calendars created once, real one shared by
membership, config written outside the repo, re-runs idempotent.
"""

import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SCRATCH = Path(tempfile.mkdtemp(prefix="crisisplan-verify32-"))
CONFIG = SCRATCH / "google.json"
STATE = SCRATCH / "gcal-state.json"

failures = []


def check(name, ok, detail=""):
    print(f"  [{'PASS' if ok else 'FAIL'}] {name}" + (f" — {detail}" if detail else ""))
    if not ok:
        failures.append(name)


def run_setup(*args, config=CONFIG):
    env = dict(os.environ,
               CRISISPLAN_GOOGLE_CONFIG=str(config),
               CRISISPLAN_GCAL_FAKE=str(STATE))
    return subprocess.run(
        [sys.executable, str(REPO / "tools" / "gsetup.py"), *args],
        capture_output=True, text=True, encoding="utf-8",
        errors="replace", env=env)


def state():
    return json.loads(STATE.read_text(encoding="utf-8"))


print("Dry run changes nothing")
r = run_setup("--dry-run")
check("dry-run exits 0", r.returncode == 0, r.stderr.strip())
check("no config or state written",
      not CONFIG.exists() and not STATE.exists())

print("First real (fake-transport) run")
r = run_setup("--emails", "alex@example.com,chris@example.com",
              "--household", "Sample family")
check("setup exits 0", r.returncode == 0, r.stderr.strip()[:200])
check("config written", CONFIG.exists())
cfg = json.loads(CONFIG.read_text(encoding="utf-8"))
check("both calendar ids recorded",
      cfg.get("test_calendar_id", "").startswith("fake-cal-")
      and cfg.get("real_calendar_id", "").startswith("fake-cal-")
      and cfg["test_calendar_id"] != cfg["real_calendar_id"])
check("credential/token paths recorded, outside the repo",
      "credentials_file" in cfg and "token_file" in cfg
      and not Path(cfg["credentials_file"]).resolve()
      .is_relative_to(REPO)
      and not Path(cfg["token_file"]).resolve().is_relative_to(REPO))
s = state()
check("exactly two calendars created", len(s["calendars"]) == 2,
      str(list(s["calendars"])))
names = [c["summary"] for c in s["calendars"].values()]
check("test calendar clearly named private",
      any("TEST" in n and "never shared" in n for n in names), str(names))
check("real calendar named for the household",
      any("Sample family" in n for n in names), str(names))
real_id = cfg["real_calendar_id"]
acl = s["acls"].get(real_id, [])
check("real calendar shared with both emails (membership)",
      sorted(a["scope"]["value"] for a in acl)
      == ["alex@example.com", "chris@example.com"]
      and all(a["role"] == "writer" for a in acl))
check("test calendar shared with nobody",
      not s["acls"].get(cfg["test_calendar_id"]))
check("both calendars (re)inserted into the owner's visible list",
      sorted(s.get("calendarList", []))
      == sorted([cfg["test_calendar_id"], cfg["real_calendar_id"]]))

print("Unsubscribed-but-not-deleted calendar is re-listed on re-run")
s["calendarList"] = []          # owner hid both calendars in the UI
STATE.write_text(json.dumps(s), encoding="utf-8")
r = run_setup("--emails", "", "--household", "Sample family")
check("re-run restores list visibility, ids unchanged",
      r.returncode == 0
      and sorted(state().get("calendarList", []))
      == sorted([cfg["test_calendar_id"], cfg["real_calendar_id"]]))

print("Re-run is idempotent (repair, never duplicate)")
r = run_setup("--emails", "alex@example.com,chris@example.com",
              "--household", "Sample family")
check("second run exits 0", r.returncode == 0)
cfg2 = json.loads(CONFIG.read_text(encoding="utf-8"))
s2 = state()
check("same calendar ids kept",
      cfg2["test_calendar_id"] == cfg["test_calendar_id"]
      and cfg2["real_calendar_id"] == cfg["real_calendar_id"])
check("still exactly two calendars", len(s2["calendars"]) == 2)
check("shares upserted, not duplicated",
      len(s2["acls"][real_id]) == 2)

print("Missing calendar is recreated (repair path)")
s2["calendars"].pop(cfg["test_calendar_id"])
STATE.write_text(json.dumps(s2), encoding="utf-8")
r = run_setup("--emails", "", "--household", "Sample family")
cfg3 = json.loads(CONFIG.read_text(encoding="utf-8"))
check("new test calendar id issued, real kept",
      cfg3["test_calendar_id"] != cfg["test_calendar_id"]
      and cfg3["real_calendar_id"] == cfg["real_calendar_id"],
      cfg3.get("test_calendar_id", "?"))

print("Guard: config inside the repo refused")
r = run_setup("--dry-run", config=REPO / "private" / "g.json")
check("refused with REFUSED", r.returncode != 0
      and "REFUSED" in (r.stderr + r.stdout))

print()
if failures:
    raise SystemExit(f"FAILED: {failures}")
print("All step-3.2 checks passed.")
