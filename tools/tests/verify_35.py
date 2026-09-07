"""Regression suite from build step 3.5: durability hardening
(AC-D5 + the annual-review self-canary from SPEC §8).
"""

import json
import os
import subprocess
import sys
import tempfile
from datetime import date, timedelta
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SAMPLE = REPO / "samples" / "pat.sample.xlsx"
SCRATCH = Path(tempfile.mkdtemp(prefix="crisisplan-verify35-"))
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
    return subprocess.run(
        [sys.executable, str(REPO / "tools" / script), *map(str, args)],
        capture_output=True, text=True, encoding="utf-8",
        errors="replace", env=ENV)


print("Annual-review self-canary (SPEC §8)")
run("gsetup.py", "--emails", "", "--household", "Fam")
cfg = json.loads(CONFIG.read_text(encoding="utf-8"))
r = run("generate.py", "-f", SAMPLE, "--date", "2030-01-15",
        "--out", SCRATCH / "run-a", "--mode", "test")
check("test push OK", r.returncode == 0 and "72 created" in r.stdout,
      r.stdout.splitlines()[-3] if r.returncode == 0 else r.stderr[:200])
events = json.loads(STATE.read_text(encoding="utf-8"))["events"][
    cfg["test_calendar_id"]]
annual = [e for e in events.values()
          if e.get("extendedProperties", {}).get("private", {})
          .get("planKey") == "meta/annual-review"]
check("exactly one annual-review event", len(annual) == 1)
expected = (date.today() + timedelta(days=365)).isoformat()
check("scheduled one year from the run date",
      annual and annual[0]["start"]["date"] == expected,
      annual[0]["start"]["date"] if annual else "missing")
check("carries the re-run instruction",
      annual and "re-run" in annual[0]["summary"])
r = run("generate.py", "-f", SAMPLE, "--date", "2030-01-15",
        "--out", SCRATCH / "run-b", "--mode", "test")
check("second run: canary skipped, not duplicated",
      "0 created" in r.stdout and len(
          [e for e in json.loads(STATE.read_text(encoding="utf-8"))
           ["events"][cfg["test_calendar_id"]].values()
           if e.get("extendedProperties", {}).get("private", {})
           .get("planKey") == "meta/annual-review"]) == 1)

print("Canary is Google-only; local artifacts stay deterministic")
md = (SCRATCH / "run-a" / "checklist.md").read_text(encoding="utf-8")
ics = (SCRATCH / "run-a" / "plan.ics").read_bytes().decode("utf-8")
preview = (SCRATCH / "run-a" / "google-preview.md").read_text(
    encoding="utf-8")
check("no annual-review event in checklist or .ics",
      "annual-review" not in md and "annual-review" not in ics)
check("preview documents the canary without listing it as an event",
      "meta/annual-review" in preview
      and "  - Plan key: meta/annual-review" not in preview)

print("AC-D5: pinned + vendored offline install")
req = (REPO / "tools" / "requirements.txt").read_text(encoding="utf-8")
check("every requirement pinned with ==",
      all("==" in line for line in req.splitlines()
          if line.strip() and not line.strip().startswith("#")))
wheels = sorted(p.name for p in (REPO / "tools" / "vendor").glob("*.whl"))
check("wheels vendored for every pinned dep",
      any(w.startswith("openpyxl-3.1.5") for w in wheels)
      and any(w.startswith("et_xmlfile-2.0.0") for w in wheels),
      ", ".join(wheels))
check("wheels are pure-Python (survive machine changes)",
      all("none-any" in w for w in wheels))
check("offline install path documented",
      "--no-index" in req and "--find-links" in req
      and "--no-index" in (REPO / "tools" / "README.md")
      .read_text(encoding="utf-8"))
r = subprocess.run(
    [sys.executable, "-m", "pip", "install", "--dry-run", "--no-index",
     "--find-links", str(REPO / "tools" / "vendor"), "-r",
     str(REPO / "tools" / "requirements.txt")],
    capture_output=True, text=True)
check("pip resolves the requirements offline (dry run)",
      r.returncode == 0, (r.stderr or r.stdout).strip()[:200])

print("Health summary covers every stage")
stdout = run("generate.py", "-f", SAMPLE, "--date", "2030-01-15",
             "--out", SCRATCH / "run-c", "--mode", "test").stdout
for stage in ("read workbook", "derive tasks", "augment cache",
              "local outputs", "friends artifact", "google config",
              "google preview", "google", "overall"):
    check(f"stage reported: {stage}", stage in stdout)

print()
if failures:
    raise SystemExit(f"FAILED: {failures}")
print("All step-3.5 checks passed.")
