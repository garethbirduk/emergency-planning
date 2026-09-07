"""Regression suite from build step 2.2: visibility layers (AC-V1–V3).

Sample rows under test (set in make_template.py):
  banks #1  (Example Bank plc — Current account)      full
  building_societies #1 (Sampletown BS — saver)       generic + ai Y
  property #1 (Family home)                           none
  everything else                                     blank -> generic
"""

import subprocess
import sys
import tempfile
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SCRATCH = Path(tempfile.mkdtemp(prefix="crisisplan-verify22-"))
sys.path.insert(0, str(REPO / "tools"))

failures = []


def check(name, ok, detail=""):
    print(f"  [{'PASS' if ok else 'FAIL'}] {name}" + (f" — {detail}" if detail else ""))
    if not ok:
        failures.append(name)


out = SCRATCH / "run-22"
r = subprocess.run(
    [sys.executable, str(REPO / "tools" / "generate.py"),
     "-f", str(REPO / "samples" / "pat.sample.xlsx"),
     "--date", "2030-01-15", "--out", str(out)],
    capture_output=True, text=True)
if r.returncode != 0:
    print(r.stdout, r.stderr)
    raise SystemExit("CLI failed")

ics = (out / "plan.ics").read_bytes().decode("utf-8").replace("\r\n ", "")
md = (out / "checklist.md").read_text(encoding="utf-8")


def events(text):
    evs = []
    for block in text.split("BEGIN:VEVENT")[1:]:
        ev = {}
        for line in block.split("\r\n"):
            if ":" in line:
                k, v = line.split(":", 1)
                ev[k.split(";")[0]] = v
        evs.append(ev)
    return evs


evs = events(ics)
by_uid = {e["UID"]: e for e in evs}

print("AC-V1/V2: generic row (building_societies #1)")
check("no 'Sampletown' anywhere in the .ics", "Sampletown" not in ics)
check("no 'sampletown' in any UID (name-derived slug gone)",
      "sampletown" not in ics.lower())
bs_notify = by_uid.get("building_societies/1/notify")
check("UID is building_societies/1/notify", bs_notify is not None)
check("summary is the generic label",
      bs_notify and bs_notify["SUMMARY"] == "Notify building society #1",
      bs_notify["SUMMARY"] if bs_notify else "missing")
check("description points at the checklist, no pointer leak",
      bs_notify and "Passbook" not in bs_notify["DESCRIPTION"]
      and "checklist" in bs_notify["DESCRIPTION"])
check("other generic rows redacted too (no rental-flat insurer name)",
      "Example Insurance Co" not in ics and "Rental flat" not in ics)

print("AC-V1: full row (banks #1)")
bank_notify = by_uid.get("banks/example-bank-plc-current-account/notify")
check("full row keeps its name-derived UID", bank_notify is not None)
check("full summary carries real name + id",
      bank_notify and "Example Bank plc" in bank_notify["SUMMARY"]
      and "(bank #1)" in bank_notify["SUMMARY"],
      bank_notify["SUMMARY"] if bank_notify else "missing")
check("full description keeps details",
      bank_notify and "survivor's" in bank_notify["DESCRIPTION"])

print("AC-V1: none row (property #1, Family home)")
check("no Family home / address anywhere in the .ics",
      "Family home" not in ics and "Example Street" not in ics)
check("no property #1 events at all",
      not any(u.startswith("property/1/") for u in by_uid))
check("rental flat (generic) events still present",
      any(u.startswith("property/2/") for u in by_uid))
check("Family home tasks still in the checklist, marked local-only",
      "Family home" in md and "local-only — not sent to any calendar" in md)

print("AC-V3: checklist legend")
check("legend decodes building_societies #1",
      "`building_societies #1` = Sampletown Building Society — "
      "Instant access saver" in md)
check("legend marks property #1 as local-only",
      "`property #1` = Family home *(local-only — no calendar entries)*"
      in md)
check("checklist keeps full detail everywhere",
      "Sampletown Building Society" in md and "Passbook" in md)
check("death-cert steps marked 📜 with a legend; icon never in event "
      "titles",
      "- [ ] 📜 " in md and "certified copy of the death certificate" in md
      and "SUMMARY:📜" not in ics)

print("Defaults for a pre-2.2 workbook (no visibility columns)")
from openpyxl import Workbook  # noqa: E402
from reader import read_workbook  # noqa: E402
from tasks import derive_tasks  # noqa: E402

old = SCRATCH / "old-style.xlsx"
wb = Workbook()
ws = wb.active
ws.title = "banks"
ws.append(["bank", "account_label", "joint?", "owner",
           "location_of_details", "notes"])
ws.append(["Legacy Bank", "Old account", "joint", "", "Desk drawer", ""])
wb.save(old)

tasks, notes = derive_tasks(read_workbook(old), date(2030, 1, 15))
legacy = [t for t in tasks if t.module == "banks"]
check("row read with defaults: generic visibility",
      legacy and all(t.visibility == "generic" for t in legacy))
check("opaque hash ref used and warned about",
      any("no id set" in n for n in notes)
      and legacy[0].external_key.startswith("banks/")
      and "legacy" not in legacy[0].external_key,
      legacy[0].external_key if legacy else "none")
check("no real name in legacy external title",
      all("Legacy" not in t.external_title for t in legacy))

print()
if failures:
    raise SystemExit(f"FAILED: {failures}")
print(f"All step-2.2 checks passed ({len(evs)} calendar events; "
      f"{md.count('local-only — not sent')} local-only task(s)).")
