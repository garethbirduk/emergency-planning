"""Browser-demo glue: the wizard + generator running under Pyodide.

Replaces wizard.py's local HTTP server with direct function calls. The
page's JS routes the wizard form's requests here:

    /save         -> save(urlencoded)   (same JSON the real server sends)
    /create       -> create(urlencoded) (writes workbook, runs the
                                         generator, returns the artifacts)
    /?load=<path> -> form_html(path)    (re-render prefilled for editing)

Everything happens in the Pyodide in-memory filesystem; nothing leaves
the browser. The real privacy guards still apply: the repo tree lives at
/app, workbooks under /data, and the guards refuse anything else.
"""

import base64
import io
import json
import sys
import urllib.parse
from contextlib import redirect_stderr, redirect_stdout
from datetime import date
from pathlib import Path

sys.path.insert(0, "/app/tools")

import wizard  # noqa: E402
from wizard import (assign_ids, build_rows, check_target,  # noqa: E402
                    ensure_planholder, load_prefill, render_form,
                    target_path, write_workbook)

DATA = Path("/data")
DEMO_TRIGGER = "2030-01-15"
_backed_up = set()


def _patch_page(page: str) -> str:
    """Adapt the wizard page for the iframe: reloads go through the
    parent (there is no server to navigate to), and the two controls
    that only make sense on a real machine are hidden."""
    page = page.replace(
        "location = '/?load=' + encodeURIComponent(data.path);",
        "parent.demoReload(data.path);")
    page = page.replace(
        "if (path) location = '/?load=' + encodeURIComponent(path);",
        "if (path) parent.demoReload(path);")
    # No native file picker and no Google setup in a browser demo.
    page = page.replace("</style>", """
button[onclick="pickLoad()"] { display: none; }
input[name="google_setup"] { pointer-events: none; }
</style>""")
    return page


def setup() -> str:
    """Copy the committed sample outside the repo tree (the wizard
    refuses to edit anything inside it) and render the prefilled form."""
    DATA.mkdir(exist_ok=True)
    demo_wb = DATA / "pat-sample" / "pat-sample.xlsx"
    demo_wb.parent.mkdir(exist_ok=True)
    demo_wb.write_bytes(
        Path("/app/samples/pat.sample.xlsx").read_bytes())
    return form_html(str(demo_wb))


def form_html(path: str) -> str:
    error, plan_for, editing, prefill, maxids = load_prefill(path)
    if error:
        return f"<p>{error}</p>"
    return _patch_page(render_form(editing=editing,
                                   plan_for_value=plan_for,
                                   prefill=prefill, maxids=maxids))


def _prepare(form: dict, kept_indexes=None):
    """wizard.WizardHandler._prepare, minus the socket."""
    plan_for = form.get("plan_for", [""])[0].strip()
    out_dir = form.get("out_dir", [""])[0].strip()
    editing = form.get("editing", [""])[0].strip()
    if not plan_for or (not editing and not out_dir):
        return (None, None, "", "The person's name and the folder are "
                "required — go back and fill them in.")
    if editing:
        path = Path(editing).resolve()
        if path.is_relative_to(wizard.REPO_ROOT):
            return (None, None, editing, "REFUSED: cannot edit a workbook "
                    "inside the repo working tree.")
        if not path.is_file():
            return (None, None, editing,
                    f"{path} no longer exists — nothing to edit.")
    else:
        path = target_path(out_dir, plan_for)
        error = check_target(path)
        if error:
            return (None, None, editing, error)
    rows_by_tab = build_rows(form, kept_indexes)
    error = ensure_planholder(rows_by_tab, plan_for)
    if error:
        return (None, None, editing, error)
    assign_ids(rows_by_tab, form)
    return (path, rows_by_tab, editing, None)


def _write(path: Path, rows_by_tab: dict, editing: str):
    """wizard.WizardHandler._write with a module-level session set."""
    backup = None
    if editing and path.is_file() and str(path) not in _backed_up:
        from datetime import datetime
        import shutil
        backup_dir = path.parent / "backups"
        backup_dir.mkdir(exist_ok=True)
        backup = backup_dir / (
            f"{path.stem}-backup-"
            f"{datetime.now().strftime('%Y%m%d-%H%M%S')}.xlsx")
        shutil.copyfile(path, backup)
        _backed_up.add(str(path))
    write_workbook(path, rows_by_tab)
    return backup


def save(body: str) -> str:
    """POST /save — same JSON contract as the real server."""
    from datetime import datetime
    form = urllib.parse.parse_qs(body)
    kept_indexes = {}
    path, rows_by_tab, editing, error = _prepare(form, kept_indexes)
    if error:
        return json.dumps({"ok": False, "error": error})
    inserted = len(rows_by_tab.get("household", [])) > len(
        kept_indexes.get("household", []))
    backup = _write(path, rows_by_tab, editing)
    ids = {}
    for tab, rows in rows_by_tab.items():
        if tab in kept_indexes and not (tab == "household" and inserted):
            ids[tab] = {str(i): (r[-3] or "")
                        for i, r in zip(kept_indexes[tab], rows)}
    return json.dumps({
        "ok": True,
        "path": str(path),
        "reload": inserted,
        "saved_at": datetime.now().strftime("%H:%M:%S"),
        "backup": backup.name if backup else "",
        "ids": ids,
        "maxids": {tab: max([int(r[-3]) for r in rows
                             if r[-3] and str(r[-3]).isdigit()],
                            default=0)
                   for tab, rows in rows_by_tab.items()},
    })


def create(body: str) -> str:
    """POST /create — write the workbook, then run the real generator
    (local mode) on it and hand every artifact back to the page."""
    form = urllib.parse.parse_qs(body)
    path, rows_by_tab, editing, error = _prepare(form)
    if error:
        return json.dumps({"ok": False, "error": error})
    _write(path, rows_by_tab, editing)

    import generate
    out = io.StringIO()
    argv, sys.argv = sys.argv, ["generate.py", "-f", str(path),
                                "--date", DEMO_TRIGGER, "--mode", "local"]
    try:
        with redirect_stdout(out), redirect_stderr(out):
            try:
                generate.main()
                code = 0
            except SystemExit as e:
                code = e.code or 0
    finally:
        sys.argv = argv

    files = {}
    for p in sorted(path.parent.iterdir()):
        if p.is_file() and p.suffix in (".md", ".ics", ".txt"):
            files[p.name] = p.read_text(encoding="utf-8")
    return json.dumps({
        "ok": code == 0,
        "path": str(path),
        "workbook": path.name,
        "workbook_b64": base64.b64encode(path.read_bytes()).decode(),
        "health": out.getvalue(),
        "trigger": DEMO_TRIGGER,
        "files": files,
    })
