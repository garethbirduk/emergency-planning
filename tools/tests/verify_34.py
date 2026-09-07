"""Regression suite from build step 3.4: real mode, --fresh, graceful
degradation (AC-M3, M4, D3).
"""

import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SAMPLE = REPO / "samples" / "pat.sample.xlsx"
SCRATCH = Path(tempfile.mkdtemp(prefix="crisisplan-verify34-"))
CONFIG = SCRATCH / "google.json"
STATE = SCRATCH / "gcal-state.json"

failures = []


def check(name, ok, detail=""):
    print(f"  [{'PASS' if ok else 'FAIL'}] {name}" + (f" — {detail}" if detail else ""))
    if not ok:
        failures.append(name)


def run(script, *args, fake=True, config=CONFIG):
    env = dict(os.environ, CRISISPLAN_GOOGLE_CONFIG=str(config))
    env.pop("CRISISPLAN_GCAL_FAKE", None)
    if fake:
        env["CRISISPLAN_GCAL_FAKE"] = str(STATE)
    return subprocess.run(
        [sys.executable, str(REPO / "tools" / script), *map(str, args)],
        capture_output=True, text=True, encoding="utf-8",
        errors="replace", env=env)


def state():
    return json.loads(STATE.read_text(encoding="utf-8"))


run("gsetup.py", "--emails", "alex@example.com", "--household", "Fam")
cfg = json.loads(CONFIG.read_text(encoding="utf-8"))

print("AC-M3: real mode writes to the configured real calendar")
r = run("generate.py", "-f", SAMPLE, "--date", "2030-01-15",
        "--out", SCRATCH / "run-real", "--mode", "real")
# 71 plan events + the meta/annual-review self-canary (step 3.5) = 72.
check("real run OK", r.returncode == 0 and "real calendar: 72 created"
      in r.stdout)
s = state()
check("events landed in the real calendar only",
      len(s["events"].get(cfg["real_calendar_id"], {})) == 72
      and not s["events"].get(cfg["test_calendar_id"]))
writes = [x for x in s["log"] if "/events" in x["path"]
          and x["method"] in ("POST", "PATCH")]
check("notifications suppressed in real mode too",
      all(x["params"].get("sendUpdates") == "none" for x in writes))

print("AC-M4: --fresh creates a disposable, clearly-named calendar")
r = run("generate.py", "-f", SAMPLE, "--date", "2030-01-15",
        "--out", SCRATCH / "run-fresh", "--mode", "test", "--fresh")
check("fresh run OK", r.returncode == 0
      and "fresh disposable calendar" in r.stdout)
s = state()
fresh_cals = [c for c in s["calendars"].values()
              if "DELETE ME" in c["summary"] and "⚠" in c["summary"]]
check("disposable calendar exists, screaming its name",
      len(fresh_cals) == 1, str([c["summary"] for c in fresh_cals]))
check("72 events in the disposable calendar",
      len(s["events"].get(fresh_cals[0]["id"], {})) == 72)
check("persistent test calendar still untouched",
      not s["events"].get(cfg["test_calendar_id"]))

print("AC-D3/M3: Google outage degrades gracefully")
s["fail"] = True
STATE.write_text(json.dumps(s), encoding="utf-8")
out_dir = SCRATCH / "run-outage"
r = run("generate.py", "-f", SAMPLE, "--date", "2030-01-15",
        "--out", out_dir, "--mode", "real")
check("run still succeeds (exit 0) with a WARN",
      r.returncode == 0 and "google ................ WARN" in r.stdout
      and "simulated Google outage" in r.stdout, str(r.returncode))
check("all local artifacts written despite the outage",
      all((out_dir / f).exists() for f in
          ("checklist.md", "plan.ics", "funeral-details.md",
           "google-preview.md", "run-info.txt")))
s = state()
s["fail"] = False
STATE.write_text(json.dumps(s), encoding="utf-8")

print("Unconfigured test/real mode: helpful WARN, no crash")
r = run("generate.py", "-f", SAMPLE, "--date", "2030-01-15",
        "--out", SCRATCH / "run-nocfg", "--mode", "test",
        fake=False, config=SCRATCH / "missing.json")
check("WARN points at gsetup",
      r.returncode == 0 and "WARN" in r.stdout
      and "gsetup" in r.stdout)
check("local artifacts complete",
      (SCRATCH / "run-nocfg" / "checklist.md").exists())

print()
if failures:
    raise SystemExit(f"FAILED: {failures}")
print("All step-3.4 checks passed.")
