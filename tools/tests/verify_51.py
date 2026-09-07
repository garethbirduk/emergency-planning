"""Regression suite from build step 5.1: the onboarding wizard.

Drives the real local server over HTTP: form renders from the template
definitions, answers become a valid workbook (outside the repo, write-
once), and the engine derives a correct plan from it.
"""

import sys
import tempfile
import threading
import urllib.parse
import urllib.request
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SCRATCH = Path(tempfile.mkdtemp(prefix="crisisplan-verify51-")).resolve()
sys.path.insert(0, str(REPO / "tools"))

import wizard  # noqa: E402
from reader import read_workbook  # noqa: E402
from tasks import derive_tasks  # noqa: E402

failures = []


def check(name, ok, detail=""):
    print(f"  [{'PASS' if ok else 'FAIL'}] {name}" + (f" — {detail}" if detail else ""))
    if not ok:
        failures.append(name)


server = wizard.make_server(0)
thread = threading.Thread(target=server.serve_forever, daemon=True)
thread.start()
BASE = f"http://127.0.0.1:{server.server_address[1]}"


def get(path="/"):
    with urllib.request.urlopen(BASE + path, timeout=10) as resp:
        return resp.status, resp.read().decode("utf-8")


def post(fields):
    data = urllib.parse.urlencode(fields).encode("utf-8")
    req = urllib.request.Request(BASE + "/create", data=data)
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.status, resp.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode("utf-8")


print("Form generated from the template definitions")
status, page = get()
check("form served", status == 200 and "Crisis-plan wizard" in page)
check("every crisis tab present as a section",
      all(f"<b>{tab}</b>" in page for tab in wizard.TAB_ORDER
          if tab != "preparation"))
check("pre-crisis / crisis panes with the checklist machinery",
      'id="pane_pre"' in page and 'id="pane_crisis"' in page
      and "showPane('crisis')" in page
      and '"Arrange cleaner"' in page      # seeded catalog
      and '"Write or update the will"' in page
      and "no surprises later" in page     # will sub-steps
      and '"Health & medical"' in page
      and "record WHERE, not the list itself" in page
      and '"Security & digital"' in page
      and '"Professional services"' in page
      and '"Family & communication"' in page
      and "password manager" in page       # pointer-shaped wording
      and "prepCloseGroup" in page and "not_needed" in page
      and "prepRestoreGroup" in page       # group-level restore
      and "seeded.has(" in page            # catalog merge on edit
      and 'placeholder="notes / outcome"' in page)
check("Enter never submits; adds items in the checklist boxes",
      "e.key !== 'Enter'" in page and "e.preventDefault()" in page
      and "prepAddCustom(e.target" in page)
check("fields come from the template columns",
      'banks__IDX__bank' in page and 'pensions__IDX__provider' in page
      and 'household__IDX__crisis_roles' in page)
check("dropdown options come from the template",
      "death_in_service" in page and "tenants_in_common" in page)
check("per-row consent checkboxes offered (except household)",
      'name="banks__IDX__online_visibility" value="full"' in page
      and 'name="banks__IDX__ai_visible?" value="Y"' in page
      and 'household__IDX__online_visibility' not in page
      and 'preparation__IDX__online_visibility' not in page)
check("AI consent only where the lookup exists, naming what it sends",
      "send this row's\n <b>bank</b> name" in page
      and 'wishes__IDX__ai_visible?' not in page
      and 'key_documents__IDX__ai_visible?' not in page
      and 'wishes__IDX__online_visibility' in page)
check("pointers-never-secrets warning shown",
      "never a password" in page)
# The browser JS stamps row indexes into the template names; the server
# parses tab__N__column. These two live in different languages, so pin
# the contract: the exact substitution the page ships, applied to a
# template name, must produce something build_rows accepts. (A real
# browser once submitted 'banks0bank' because of a bad replaceAll —
# every typed row was silently dropped.)
check("JS row-index substitution keeps the __N__ separators",
      "replaceAll('__IDX__', '__' + idx + '__')" in page)
check("plan-for name live-mirrors into the first household row",
      "syncPlanholder" in page and "dataset.dirty" in page
      and "household__0__name" in page)
check("sections start empty; only household gets a starter row",
      "addRow('household');" in page
      and "tpl.id.slice(4)" not in page)
check("row counts + scaffold rows wired",
      'id="count_banks">[0]<' in page
      and "addRow('banks', true)" in page
      and "updateCount" in page and "removeRow(" in page
      and "tab + '_' + (idx + 1)" in page)
check("live calendar-label preview wired per row",
      'class="gpreview"' in page and "refreshPreview" in page
      and '"banks": ["bank", "account_label"]' in page
      and '"banks": "bank"' in page)
simulated = "banks__IDX__bank".replace("__IDX__", "__0__")
check("a JS-transformed name parses into a row",
      wizard.build_rows({simulated: ["Probe Bank"]})
      .get("banks", [("",)])[0][0] == "Probe Bank")
# Emulate a browser end to end: every field name the page's templates
# would render for row 0 must be either parseable into that tab's row
# or a recognised consent checkbox.
import re  # noqa: E402

for tab, inner in re.findall(r'<template id="tpl_(\w+)">(.*?)</template>',
                             page, re.S):
    field_names = re.findall(r'name="([^"]+)"',
                             inner.replace("__IDX__", "__0__"))
    parsed = wizard.build_rows(
        {name: ["x"] for name in field_names})
    if not (len(parsed.get(tab, [])) == 1):
        check(f"every rendered field lands in the {tab} row", False,
              str(field_names))
        break
else:
    check("every rendered template field round-trips to its tab's row",
          True)

print("Native picker endpoint")
import os  # noqa: E402

check("open button wired on the form",
      "pickLoad()" in page and "/pick?kind=" in page
      and "pickFolder()" not in page)
try:
    status, _ = get("/pick?kind=bogus")
except urllib.error.HTTPError as e:
    status = e.code
check("unknown picker kind rejected", status == 400)
check("picker opens at the nearest existing ancestor of the field",
      wizard.initial_dir(str(SCRATCH / "does" / "not" / "exist"))
      == str(SCRATCH)
      and Path(wizard.initial_dir("")).is_dir())
os.environ["CRISISPLAN_PICK_FAKE"] = str(SCRATCH / "picked.xlsx")
try:
    status, picked = get("/pick?kind=workbook")
    check("picker returns the chosen absolute path (fake hook)",
          status == 200 and picked == str(SCRATCH / "picked.xlsx"))
finally:
    del os.environ["CRISISPLAN_PICK_FAKE"]
# The dialog opened BEHIND the browser until the Tk root was pumped
# through an event loop first — and with a single-threaded server the
# page froze meanwhile, so the picker looked like it had vanished.
picker_src = (REPO / "tools" / "wizard.py").read_text(encoding="utf-8")
dialog_setup = picker_src.split("import tkinter as tk")[1].split(
    "askopenfilename")[0]
check("the picker is brought to the front before it opens",
      all(call in dialog_setup for call in
          ("root.update()", "root.lift()", "root.focus_force()")))
check("a picker that can't open says so instead of doing nothing",
      'chosen = f"ERROR: {type(exc).__name__}: {exc}"' in picker_src
      and "text.startsWith('ERROR:')" in page)
check("the page says a native window is waiting while it's open",
      'id="pickmsg"' in page and 'class="pickbtn"' in page
      and "pickBusy(true" in page)

print("Guards")
common = {"plan_for": "Sam Example",
          "out_dir": str(SCRATCH / "out")}
status, page = post({**common, "plan_for": ""})
check("missing name rejected", status == 400)
status, page = post({**common, "out_dir": str(REPO / "private")})
check("repo-tree destination refused",
      status == 400 and "REFUSED" in page)
status, page = post({**common,
                     "household__0__name": "Somebody Else",
                     "household__0__role": "planholder"})
check("planholder row naming a different person rejected",
      status == 400 and "make them match" in page)
status, page = post({**common,
                     "household__0__name": "Sam Example",
                     "household__0__role": "planholder",
                     "household__1__name": "Val Example",
                     "household__1__role": "planholder"})
check("two planholder rows rejected",
      status == 400 and "exactly one person" in page)

print("A filled interview becomes a working plan")
answers = {
    **common,
    "household__0__name": "Sam Example",
    "household__0__relationship": "self",
    "household__0__role": "planholder",
    "household__1__name": "Val Example",
    "household__1__relationship": "spouse",
    "household__1__role": "executor",
    "household__1__crisis_roles": "admin, organiser, comms",
    "household__3__name": "Kim Example",
    "household__3__relationship": "sibling",
    "household__3__role": "next_of_kin",
    "household__3__crisis_roles": "local, medical",
    "banks__0__bank": "Wizard Bank plc",
    "banks__0__account_label": "Current account",
    "banks__0__joint?": "joint",
    "banks__0__online_visibility": "full",   # ticked checkboxes
    "banks__0__ai_visible?": "Y",
    "banks__2__bank": "Second Bank",       # sparse indexes are fine
    "banks__2__account_label": "Saver",
    "banks__2__joint?": "sole",
    "banks__2__owner": "Sam",
    "wishes__0__topic": "burial_or_cremation",
    "wishes__0__wish": "Cremation",
    "life_insurance__0__provider": "Wizard Employer Ltd",
    "life_insurance__0__policy_label": "Death in service",
    "life_insurance__0__type": "death_in_service",
    "life_insurance__0__owner": "Sam",
    "life_insurance__0__nomination_in_place?": "Y",
}
status, page = post(answers)
check("submission succeeds", status == 200 and "Workbook written" in page,
      page[:200] if status != 200 else "")
out_path = SCRATCH / "out" / "sam-example" / "sam-example.xlsx"
check("workbook written into a per-person subfolder (slugified name)",
      out_path.exists())
check("per-person subfolder appended once, never doubled",
      wizard.target_path(str(SCRATCH / "out"), "Sam Example")
      == out_path
      and wizard.target_path(str(SCRATCH / "out" / "sam-example"),
                             "Sam Example") == out_path)
check("next-steps include the exact generate command",
      "generate.py" in page and str(out_path) in page)

print("First plan generated and linked on the success page")
check("plan files written alongside the workbook",
      (out_path.parent / "checklist.md").exists()
      and (out_path.parent / "google-preview.md").exists())
check("links to checklist + google preview",
      "/file?p=" in page and ">checklist.md</a>" in page
      and ">google-preview.md</a>" in page
      and ">checklist-val-example.md</a>" in page)
link = re.search(r'href="(/file\?p=[^"]+google-preview\.md)"', page)
status, preview = get(link.group(1).replace("&amp;", "&"))
check("linked preview is readable and redacted by the checkboxes",
      status == 200 and "Notify Wizard Bank plc" in preview
      and "Second Bank" not in preview)
try:
    status, _ = get("/file?p=" + urllib.parse.quote(
        str(REPO / "tools" / "gconfig.py")))
except urllib.error.HTTPError as e:
    status = e.code
check("file serving refuses anything outside the run folder",
      status == 403)
try:
    status, _ = get("/file?p=" + urllib.parse.quote(str(out_path)))
except urllib.error.HTTPError as e:
    status = e.code
check("file serving never serves the workbook itself", status == 403)
status, done_page = get("/done")
check("Finish link closes the wizard", status == 200
      and "close this window" in done_page)

model = read_workbook(out_path)
check("reader accepts the wizard workbook",
      set(model.modules) == {"household", "banks", "wishes",
                             "life_insurance"},
      ", ".join(model.modules))
check("empty tabs read as not-applicable",
      "cash_isas" in model.empty_tabs and "funeral" in model.empty_tabs)
ticked = model.modules["banks"][0]
check("ticked checkboxes -> full visibility + AI consent on that row",
      ticked.visibility == "full" and ticked.ai_visible)
check("unticked rows keep the safe defaults",
      all(r.visibility == "generic" and not r.ai_visible
          for tab, rows in model.modules.items() for r in rows
          if not (tab == "banks" and r.row_id == "1")))
check("ids pre-filled by the template",
      model.modules["banks"][0].row_id == "1")

tasks, notes = derive_tasks(model, date(2030, 6, 1))
keys = {t.plan_key for t in tasks}
check("engine derives a sensible plan from it",
      any(k.startswith("banks/wizard-bank-plc") for k in keys)
      and any(k.endswith("/claim-benefits") for k in keys)
      and all(t.assignee in ("Val Example", "Kim Example")
              for t in tasks)
      and any(t.assignee == "Kim Example" for t in tasks))
check("joint vs sole branching applied to the answers",
      any(k.endswith("wizard-bank-plc-current-account/confirm-survivor")
          for k in keys)
      and any(k.endswith("second-bank-saver/dod-balance") for k in keys))

print("Save on demand: the form stays open, ids stay put")
import json  # noqa: E402

# The first server shut itself down on /done; saving needs a live one.
server_s = wizard.make_server(0)
threading.Thread(target=server_s.serve_forever, daemon=True).start()
BASE = f"http://127.0.0.1:{server_s.server_address[1]}"


def save(fields):
    data = urllib.parse.urlencode(fields, doseq=True).encode("utf-8")
    req = urllib.request.Request(BASE + "/save", data=data)
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode("utf-8"))


mid = {"plan_for": "Mid Interview",
       "out_dir": str(SCRATCH / "midway"),
       "household__0__name": "Mid Interview",
       "household__0__role": "planholder",
       "banks__0__bank": "Halfway Bank",
       "banks__0__account_label": "Current"}
first = save(mid)
saved_path = Path(first.get("path", ""))
check("a save mid-interview writes the workbook",
      first["ok"] and saved_path.is_file(), first.get("error", ""))
check("no backup on the very first write (there was nothing to keep)",
      first["backup"] == "")
check("the ids it issued come back keyed by form row",
      first["ids"]["banks"] == {"0": "1"}
      and first["maxids"]["banks"] == 1, json.dumps(first["ids"]))
check("the page is told what to overwrite next time",
      first["path"] and not first["reload"])

# Carry on typing: same rows plus a new one, saved again.
mid2 = dict(mid, editing=first["path"],
            maxid_banks=str(first["maxids"]["banks"]),
            **{"banks__0__id": first["ids"]["banks"]["0"],
               "banks__1__bank": "Second Bank",
               "banks__1__account_label": "Saver",
               "wishes__0__topic": "burial_or_cremation",
               "wishes__0__wish": "Cremation"})
second = save(mid2)
check("saving again overwrites the same file", second["ok"]
      and second["path"] == first["path"])
check("the first row keeps its id; the new row gets the next one",
      second["ids"]["banks"] == {"0": "1", "1": "2"},
      json.dumps(second["ids"]["banks"]))
check("the version from before this session is kept, once",
      second["backup"].startswith("mid-interview-backup-"))
third = save(dict(mid2, **{"banks__1__id": "2"}))
check("later saves don't bury it in near-identical backups",
      third["ok"] and third["backup"] == "")
backups = list((saved_path.parent / "backups").glob("*.xlsx"))
check("exactly one backup on disk for the session", len(backups) == 1,
      f"{len(backups)} file(s)")

saved_model = read_workbook(saved_path)
check("the workbook holds everything typed so far",
      {"household", "banks", "wishes"} <= set(saved_model.modules)
      and len(saved_model.modules["banks"]) == 2)
check("a save with nothing to save is refused, not written",
      save({"plan_for": "", "out_dir": ""})["ok"] is False)
_status, form_page = get("/")
check("load/save live in a toolbar that stays put as the form scrolls",
      'class="toolbar"' in form_page and "position: sticky" in form_page
      and 'id="savebtn"' in form_page and "pickLoad()" in form_page
      and 'id="savemsg"' in form_page)
check("unsaved answers are guarded on the way out",
      "beforeunload" in wizard.SCRIPT_JS
      and "leaving = true" in form_page)

print("Write-once")
server2 = wizard.make_server(0)   # first server shut down after success
t2 = threading.Thread(target=server2.serve_forever, daemon=True)
t2.start()
BASE = f"http://127.0.0.1:{server2.server_address[1]}"
status, page = post(answers)
check("second run with the same name refused, file untouched",
      status == 400 and "never overwrites" in page)

print("Planholder auto-added when the person isn't listed")
status, page = post({"plan_for": "Solo Person",
                     "out_dir": str(SCRATCH / "out"),
                     "banks__0__bank": "Only Bank",
                     "banks__0__account_label": "Account",
                     "banks__0__joint?": "sole"})
check("submission succeeds without a household section", status == 200)
solo = read_workbook(SCRATCH / "out" / "solo-person"
                     / "solo-person.xlsx")
holder = solo.modules["household"][0]
check("planholder row added for the person, visibly labelled",
      holder.fields.get("name") == "Solo Person"
      and holder.fields.get("role") == "planholder"
      and "Added by the wizard" in (holder.fields.get("notes") or ""))
check("engine treats them as the deceased (sole bank -> probate path)",
      any(t.plan_key.endswith("only-bank-account/dod-balance")
          for t in derive_tasks(solo, date(2030, 6, 1))[0]))
check("two plans share one root without clobbering each other",
      (SCRATCH / "out" / "solo-person" / "checklist.md").exists()
      and "Wizard Bank plc"
      in (out_path.parent / "checklist.md").read_text(encoding="utf-8"))

print("Editor: load, rename, remove a row, add a row")
status, editpage = get("/?load=" + urllib.parse.quote(str(out_path)))
check("edit form prefilled from the workbook",
      status == 200 and "Wizard Bank plc" in editpage
      and f'name="editing" value="{out_path}"' in editpage
      and 'value="Sam Example"' in editpage)
check("remove-row button present", "remove row" in editpage)
check("max-ever id carried per tab",
      'name="maxid_banks" value="2"' in editpage)
status, page2 = post({
    "plan_for": "Sam Example",
    "editing": str(out_path),
    "maxid_banks": "3",       # pretend id 3 was used and deleted before
    "household__0__name": "Sam Example",
    "household__0__relationship": "self",
    "household__0__role": "planholder",
    "household__0__id": "1",
    "household__1__name": "Val Example",
    "household__1__role": "executor",
    "household__1__crisis_roles": "admin, organiser, comms",
    "household__1__id": "2",
    "banks__0__bank": "Renamed Bank plc",
    "banks__0__account_label": "Current account",
    "banks__0__joint?": "joint",
    "banks__0__id": "1",
    "banks__0__online_visibility": "full",
    # 'Second Bank' (id 2) not re-posted -> the row is deleted
    "banks__5__bank": "Third Bank",       # new row, no id yet
    "banks__5__account_label": "New saver",
    "banks__5__joint?": "sole",
})
check("edit save succeeds with a backup note",
      status == 200 and "backed up as" in page2, page2[:200])
backups = list((out_path.parent / "backups")
               .glob("sam-example-backup-*.xlsx"))
check("timestamped backup written under backups/", len(backups) == 1)
edited = read_workbook(out_path)
check("tabs not re-posted are removed (wishes/life gone)",
      set(edited.modules) == {"household", "banks"})
banks = {r.row_id: r for r in edited.modules["banks"]}
check("rename kept its id and visibility; deleted id gone; new row "
      "skips ever-used ids",
      banks.get("1") is not None
      and banks["1"].fields["bank"] == "Renamed Bank plc"
      and banks["1"].visibility == "full"
      and "2" not in banks and "3" not in banks
      and banks.get("4") is not None
      and banks["4"].fields["bank"] == "Third Bank")
status, page3 = post({"plan_for": "X", "editing":
                      str(REPO / "samples" / "pat.sample.xlsx")})
check("editing a repo-tree workbook refused",
      status == 400 and "REFUSED" in page3)
server2.shutdown()

check("wizard makes no network-capable imports beyond the local server",
      "urllib.request" not in (REPO / "tools" / "wizard.py")
      .read_text(encoding="utf-8"))

print()
if failures:
    raise SystemExit(f"FAILED: {failures}")
print("All step-5.1 checks passed.")
