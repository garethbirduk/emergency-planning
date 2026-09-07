#!/usr/bin/env python3
"""Preparation plan — "what to do before someone dies" (step 5.3).

Renders ``preparation.md``: the declared rows from the ``preparation``
tab (actions, decisions, living preferences like "no stairlift") plus
**gaps the tool notices** in the rest of the workbook (no will, no LPA,
missing pension nominations…). Deliberately simple (SPEC §16 first
slice): a local-only markdown file — no calendar, no dates math, no
visibility machinery. Status lives in the workbook, human-owned; a
noticed gap clears itself the moment the workbook fixes it.

Usage:  python tools/prep.py <workbook.xlsx>
"""

import sys
from pathlib import Path

from reader import slugify

PRIORITY_ORDER = {"now": 0, "soon": 1, "someday": 2, "": 3}


def _live_actions(model) -> list:
    return [r for r in model.modules.get("preparation", [])
            if (r.fields.get("status") or "").lower() != "not_needed"
            and (r.fields.get("type") or "action").lower() == "action"]


def _grouped_action_lines(actions) -> list:
    """Shared layout: ungrouped first, then group subsections, items
    sorted by priority within each."""
    lines = []
    groups = {}
    for r in actions:
        groups.setdefault(r.fields.get("group") or "", []).append(r)
    for group, items in groups.items():
        items.sort(key=lambda r: PRIORITY_ORDER.get(
            (r.fields.get("priority") or "").lower(), 3))
        lines += ["", f"### {group}", ""] if group else [""]
        for r in items:
            lines += _action_line(r)
    return lines


def derive_gaps(model) -> list:
    """Crisis-plan COMPLETENESS gaps: information the workbook is
    missing, noticed from the other tabs — distinct from the declared
    pre-crisis life work on the preparation tab. Deterministic rules;
    each self-clears when the workbook gains the fix. Judgment calls
    that depend on age/circumstances (e.g. account consolidation) are
    deliberately excluded until the plan has a basis for them
    (PLAN 5.4)."""
    gaps = []
    documents = {(r.fields.get("document") or "").lower()
                 for r in model.modules.get("key_documents", [])}
    if not any(d.startswith("will") for d in documents):
        gaps.append("No will recorded on the key_documents tab — if one "
                    "exists, record where it is; if not, make one.")
    if not any(d.startswith("lpa") for d in documents):
        gaps.append("No LPA recorded — set up finance and health LPAs "
                    "while everyone is able (they take weeks to "
                    "register).")
    if "wishes" not in model.modules:
        gaps.append("The wishes tab is empty — record "
                    "funeral/cremation/donation wishes while they can "
                    "be asked about.")
    for row in model.modules.get("pensions", []):
        if (row.fields.get("type") or "").lower() == "state":
            continue
        if (row.fields.get("nomination_in_place?") or "").upper() != "Y":
            gaps.append(f"Pension '{row.label}' has no expression of "
                        "wish — file one so benefits pay without "
                        "trustee discretion.")
    for row in model.modules.get("life_insurance", []):
        if (row.fields.get("nomination_in_place?") or "").upper() != "Y":
            gaps.append(f"Life policy '{row.label}' has no nomination / "
                        "trust noted — check with the insurer; in trust "
                        "it pays without probate.")
    if "funeral" not in model.modules:
        gaps.append("No pre-arranged funeral recorded — optional, but "
                    "it removes the biggest day-2 decision.")
    return gaps


def _action_line(r) -> list:
    status = (r.fields.get("status") or "todo").lower()
    box = "x" if status == "done" else " "
    priority = (r.fields.get("priority") or "").upper()
    prefix = f"**{priority}** " if priority else ""
    owner = r.fields.get("owner")
    tail = f" — *{owner}*" if owner else ""
    state = " *(in progress)*" if status == "in_progress" else ""
    lines = [f"- [{box}] {prefix}{r.fields.get('item') or ''}"
             f"{tail}{state}"]
    if r.fields.get("notes"):
        lines.append(f"  - {r.fields['notes']}")
    return lines


def render_preparation(model) -> str:
    rows = model.modules.get("preparation", [])
    # Closed ("not needed") items stay out of the active plan entirely;
    # a footer names them so the choice is visible, never silent.
    closed = [r for r in rows
              if (r.fields.get("status") or "").lower() == "not_needed"]
    live = [r for r in rows if r not in closed]
    actions = [r for r in live
               if (r.fields.get("type") or "action").lower() == "action"]
    other = [r for r in live
             if (r.fields.get("type") or "action").lower() != "action"]
    gaps = derive_gaps(model)

    lines = [
        f"# Preparing now — {model.source.stem}",
        "",
        "> Work to do **while everyone is alive** — the other half of "
        "the crisis plan. Track status in the spreadsheet's "
        "`preparation` tab (the tool only reads it). The completeness "
        "section at the bottom flags information the crisis plan "
        "itself is still missing.",
    ]
    if actions:
        lines += ["", "## Actions"]
        lines += _grouped_action_lines(actions)
    if other:
        lines += ["", "## Preferences & notes", ""]
        for r in other:
            note = r.fields.get("notes")
            lines.append(f"- {r.fields.get('item') or ''}"
                         + (f" — {note}" if note else ""))
    if gaps:
        lines += ["", "## Crisis-plan completeness", "",
                  "The crisis plan itself is missing information — "
                  "record it in the workbook and these vanish. (These "
                  "are about the *plan* being ready, not life admin — "
                  "that's the Actions section above.)", ""]
        lines += [f"- {g}" for g in gaps]
    if not (actions or other or gaps):
        lines += ["", "Nothing to prepare — the workbook has no "
                      "preparation rows and no noticed gaps."]
    if closed:
        names = "; ".join(
            (f"{r.fields.get('group')}: " if r.fields.get("group")
             else "") + (r.fields.get("item") or "")
            for r in closed)
        lines += ["", f"*Marked not needed: {names}.*"]
    lines.append("")
    return "\n".join(lines)


def render_person_preparation(model, owner: str) -> str:
    mine = [r for r in _live_actions(model)
            if (r.fields.get("owner") or "").strip() == owner]
    lines = [
        f"# Preparing now — {model.source.stem} — {owner}",
        "",
        f"> Your share of the preparation work — {len(mine)} action(s). "
        "The full picture (everyone's actions, preferences and the "
        "completeness check) is in preparation.md.",
    ]
    lines += _grouped_action_lines(mine)
    lines.append("")
    return "\n".join(lines)


def write_preparation(model, out_dir):
    """Write preparation.md (+ per-owner lists when the work is shared
    between more than one person); return (path, detail)."""
    rows = model.modules.get("preparation", [])
    gaps = derive_gaps(model)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / "preparation.md"
    path.write_text(render_preparation(model), encoding="utf-8")
    # Personal lists, mirroring the crisis checklists: regenerated in
    # place, so stale ones (owner renamed/removed) are dropped first.
    for stale in out_dir.glob("preparation-*.md"):
        stale.unlink()
    owners = sorted({(r.fields.get("owner") or "").strip()
                     for r in _live_actions(model)} - {""})
    if len(owners) > 1:
        for owner in owners:
            (out_dir / f"preparation-{slugify(owner)}.md").write_text(
                render_person_preparation(model, owner),
                encoding="utf-8")
    done = sum(1 for r in rows
               if (r.fields.get("status") or "").lower() == "done")
    personal = (f", {len(owners)} personal list(s)"
                if len(owners) > 1 else "")
    return path, (f"preparation.md ({len(rows)} item(s), {done} done; "
                  f"{len(gaps)} completeness gap(s){personal})")


def main() -> None:
    if len(sys.argv) != 2:
        sys.exit("usage: python tools/prep.py <workbook.xlsx>")
    from reader import read_workbook
    print(render_preparation(read_workbook(sys.argv[1])))


if __name__ == "__main__":
    main()
