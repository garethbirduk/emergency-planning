#!/usr/bin/env python3
"""Output generators: markdown checklist + .ics calendar (build step 1.4).

Renders the task engine's dated Tasks into the two always-local artifacts
(SPEC §7, AC-D1 partial):

- **Markdown checklist** — human-readable, grouped by date with day offsets
  from the trigger, checkboxes per task, and the skip notes so nothing is
  silently dropped.
- **`.ics` calendar** — RFC 5545, built with the standard library only (see
  requirements.txt for why). One all-day VEVENT per task with
  ``UID = planKey`` so calendar apps treat a re-imported plan as updates to
  the same events, not duplicates (AC-I1/I2/I3 groundwork).

Both renderers are deterministic: identical inputs give byte-identical
output. DTSTAMP is pinned to the trigger date, not "now", for that reason.

Usage:  python tools/outputs.py <workbook.xlsx> --date YYYY-MM-DD --out DIR
        (writes checklist.md and plan.ics into DIR)
"""

import argparse
import re
from datetime import date, datetime, time, timedelta
from pathlib import Path

from reader import read_workbook, slugify
from tasks import Task, derive_tasks

CRLF = "\r\n"
FOLD_LIMIT = 75  # octets per line before folding (RFC 5545 §3.1)


# --- markdown checklist ------------------------------------------------------

# Task line ends with its inline key; older files carried it as a
# "  - key:" sub-bullet — both are read so ticks survive the migration.
_INLINE_KEY = re.compile(r"^- \[(.)\] .*`([^`]+)`\s*$")
_LEGACY_KEY = re.compile(r"^  - key: `([^`]+)`")


def read_checked(out_dir) -> set:
    """Plan keys the human has ticked in the existing checklist files.
    Ticks are status, and status is human-owned (AC-I4): regeneration
    re-applies them by key, so they survive retitles, date changes and
    reordering. A tick in any checklist (master or personal) counts."""
    checked = set()
    for path in Path(out_dir).glob("checklist*.md"):
        ticked = False
        for line in path.read_text(encoding="utf-8").splitlines():
            if line.startswith("- ["):
                ticked = line[:5].lower() == "- [x]"
                inline = _INLINE_KEY.match(line)
                if inline and ticked:
                    checked.add(inline.group(2))
            legacy = _LEGACY_KEY.match(line)
            if legacy and ticked:
                checked.add(legacy.group(1))
    return checked


def _task_section(tasks, trigger: date, show_assignee: bool,
                  augment=None, checked=frozenset()) -> list:
    """The date-grouped task block shared by master and personal lists.

    ``augment`` maps a row planKey to cached lookup lines (SPEC §13);
    they merge into the row's earliest task in this list, and into the
    markdown only — the .ics is an external surface and never carries
    them (institution-specific lines would defeat a redacted row)."""
    lines = []
    merged = set()
    by_date = {}
    for t in tasks:
        by_date.setdefault(t.due, []).append(t)
    for due in sorted(by_date):
        offset = (due - trigger).days
        lines += ["", f"## {due.isoformat()} (day {offset})", ""]
        for t in by_date[due]:
            # Reads as an instruction: WHO (role, holder), then what to
            # do, then the stable key as a pointer (SPEC §14).
            who = (f"**{t.role.upper()} ({t.assignee}):**"
                   if show_assignee and t.assignee
                   else f"**{t.role.upper()}:**")
            box = "x" if t.plan_key in checked else " "
            cert = "📜 " if t.needs_cert else ""
            lines.append(f"- [{box}] {cert}{who} {t.title} → "
                         f"`{t.plan_key}`")
            for d in t.details:
                lines.append(f"  - {d}")
            row_key = t.plan_key.rsplit("/", 1)[0]
            if augment and row_key in augment and row_key not in merged:
                merged.add(row_key)
                for extra in augment[row_key]:
                    lines.append(f"  - [ai] {extra}")
            if t.visibility == "none":
                lines.append("  - local-only — not sent to any calendar")
    return lines


def render_markdown(tasks, notes, source, trigger: date,
                    augment=None, checked=frozenset()) -> str:
    """Human-readable dated checklist, grouped by due date."""
    lines = [
        f"# Crisis plan — {Path(source).stem}",
        "",
        f"- **Trigger date:** {trigger.isoformat()}",
        f"- **Source workbook:** {Path(source).name}",
        f"- **Tasks:** {len(tasks)}",
        "",
        "> Work top to bottom. Tick items off here or in the calendar. "
        "Keep a printed copy with the will.",
        ">",
        "> 📜 = have a certified copy of the death certificate ready "
        "for this step (they're returned — the register-death task "
        "says how many to order).",
    ]
    lines += _task_section(tasks, trigger, show_assignee=True,
                           augment=augment, checked=checked)
    # Legend decoding the redacted calendar references (AC-V3). The
    # checklist is the local, full-detail artifact — this is where a
    # "building_societies #1" event becomes a real name again.
    redacted = {}
    for t in tasks:
        if t.visibility != "full":
            redacted[(t.module, t.row_ref)] = (t.row_label, t.visibility)
    if redacted:
        lines += ["", "## Calendar legend — redacted rows", "",
                  "Calendar entries show only these references; "
                  "they mean:", ""]
        for (module, ref), (label, vis) in sorted(redacted.items()):
            suffix = (" *(local-only — no calendar entries)*"
                      if vis == "none" else "")
            lines.append(f"- `{module} #{ref}` = {label}{suffix}")
    if notes:
        lines += ["", "## Notes — skipped items & warnings", ""]
        lines += [f"- {n}" for n in notes]
    lines.append("")
    return "\n".join(lines)


def render_person_markdown(tasks, source, trigger: date,
                           person: str, augment=None,
                           checked=frozenset()) -> str:
    """One assignee's personal checklist (AC-R2): their tasks only, same
    grouping; the master checklist keeps the notes and legend."""
    mine = [t for t in tasks if t.assignee == person]
    lines = [
        f"# Crisis plan — {Path(source).stem} — {person}",
        "",
        f"- **Trigger date:** {trigger.isoformat()}",
        f"- **Your tasks:** {len(mine)} of {len(tasks)}",
        "",
        "> Your share of the plan. The master checklist.md has everyone's "
        "tasks, the notes, and the calendar legend.",
    ]
    lines += _task_section(mine, trigger, show_assignee=False,
                           augment=augment, checked=checked)
    lines.append("")
    return "\n".join(lines)


# --- .ics calendar -----------------------------------------------------------

def _escape(text: str) -> str:
    """Escape TEXT values per RFC 5545 §3.3.11."""
    return (text.replace("\\", "\\\\").replace(";", "\\;")
                .replace(",", "\\,").replace("\n", "\\n"))


def _fold(line: str) -> str:
    """Fold a content line at 75 octets (UTF-8 safe) per RFC 5545 §3.1."""
    raw = line.encode("utf-8")
    if len(raw) <= FOLD_LIMIT:
        return line
    parts = []
    limit = FOLD_LIMIT
    while raw:
        cut = min(limit, len(raw))
        # Never split inside a multi-byte UTF-8 sequence (continuation
        # bytes are 0b10xxxxxx).
        while cut < len(raw) and (raw[cut] & 0xC0) == 0x80:
            cut -= 1
        parts.append(raw[:cut].decode("utf-8"))
        raw = raw[cut:]
        limit = FOLD_LIMIT - 1  # continuation lines start with a space
    return (CRLF + " ").join(parts)


def _external_description(task: Task) -> str:
    """The description an external surface may carry (SPEC §13): full
    detail for `full` rows, only a pointer home for redacted rows. Used
    by both the .ics and the google-preview so they can never differ."""
    if task.visibility == "full":
        return "\n".join(task.details + [f"Plan key: {task.external_key}"])
    return "\n".join([f"See the local checklist — row #{task.row_ref}.",
                      f"Plan key: {task.external_key}"])


def _event(task: Task, trigger: date) -> list:
    dtstamp = trigger.strftime("%Y%m%dT000000Z")  # pinned: deterministic
    # The .ics is an external surface (SPEC §13): redacted tasks carry no
    # real names in any property — summary, description, or UID.
    description = _external_description(task)
    return [
        "BEGIN:VEVENT",
        f"UID:{task.external_key}",
        f"DTSTAMP:{dtstamp}",
        f"DTSTART;VALUE=DATE:{task.due.strftime('%Y%m%d')}",
        f"DTEND;VALUE=DATE:{(task.due + timedelta(days=1)).strftime('%Y%m%d')}",
        f"SUMMARY:{_escape(task.external_title)}",
        f"DESCRIPTION:{_escape(description)}",
        f"CATEGORIES:{_escape(task.module)}",
        "TRANSP:TRANSPARENT",
        "END:VEVENT",
    ]


def render_ics(tasks, trigger: date) -> str:
    """RFC 5545 calendar: one all-day event per task, UID = external key.

    Tasks with visibility ``none`` are omitted entirely — they exist only
    in the local checklist (AC-V1)."""
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//crisis-plan//generator//EN",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
    ]
    for task in tasks:
        if task.visibility == "none":
            continue
        lines += _event(task, trigger)
    lines.append("END:VCALENDAR")
    return CRLF.join(_fold(line) for line in lines) + CRLF


# --- Google publish preview (build step 3.1, SPEC §7, AC-M7) -----------------
#
# The human-readable manifest of exactly what the Google stage would
# create or update, written every run before any network stage exists or
# runs. It shows the SAME external surface as the .ics — titles,
# descriptions and keys via the same code paths — so the family can read
# what would leave the house before the first test/real push.


def render_google_preview(tasks, trigger: date, mode: str,
                          calendar_note: str) -> str:
    published = [t for t in tasks if t.visibility != "none"]
    withheld = [t for t in tasks if t.visibility == "none"]
    lines = [
        "# Google Calendar preview — what would be published",
        "",
        f"- **This run's mode:** {mode}"
        + (" — nothing was sent; this file is the preview"
           if mode == "local" else ""),
        f"- **Target calendar:** {calendar_note}",
        f"- **Events to publish:** {len(published)}",
        f"- **Withheld (online_visibility = none):** {len(withheld)} "
        "task(s)",
        "",
        "> Every event below is exactly what the shared calendar would "
        "show — titles and descriptions follow each row's "
        "online_visibility, and nothing else leaves this computer. Read "
        "this file before the first real push.",
        "",
        "*A test/real push also maintains one extra event: an "
        "'annual review — re-run the tool' reminder a year from the "
        "run (key `meta/annual-review`), the plan's self-canary.*",
    ]
    by_date = {}
    for t in published:
        by_date.setdefault(t.due, []).append(t)
    for due in sorted(by_date):
        offset = (due - trigger).days
        lines += ["", f"## {due.isoformat()} (day {offset})", ""]
        for t in by_date[due]:
            lines.append(f"- **{t.external_title}**")
            for d in _external_description(t).split("\n"):
                lines.append(f"  - {d}")
    if withheld:
        by_row = {}
        for t in withheld:
            by_row.setdefault((t.module, t.row_ref), []).append(t)
        lines += ["", "## Withheld rows — never sent", "",
                  "These stay in the local checklist only:", ""]
        for (module, ref), row_tasks in sorted(by_row.items()):
            lines.append(f"- `{module} #{ref}` — {len(row_tasks)} "
                         "task(s), local checklist only")
    lines.append("")
    return "\n".join(lines)


def write_google_preview(tasks, trigger: date, mode: str,
                         calendar_note: str, out_dir):
    """Write google-preview.md; return (path, detail) for the health
    summary."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / "google-preview.md"
    path.write_text(
        render_google_preview(tasks, trigger, mode, calendar_note),
        encoding="utf-8")
    published = sum(1 for t in tasks if t.visibility != "none")
    withheld = len(tasks) - published
    return path, (f"google-preview.md ({published} event(s); "
                  f"{withheld} withheld)")


# --- friends-facing funeral artifact (build step 2.6, SPEC §7) ---------------
#
# A distinct audience gets a distinct artifact: funeral-details.md + a
# single-event funeral-details.ics for the wider circle, built ONLY from
# the funeral tab — and only from the share fields (provider, venue,
# wake_venue, service_date, service_time), so no other tab's rows (and no
# internal pointer/package/prepaid fields) can leak into something meant
# to be forwarded. Generated once a service_date is known (the re-run
# flow); regenerated idempotently.


def _as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        try:
            return date.fromisoformat(value.strip())
        except ValueError:
            return None
    return None


def _as_time(value):
    if isinstance(value, datetime):
        return value.time()
    if isinstance(value, time):
        return value
    if isinstance(value, str):
        try:
            return time.fromisoformat(value.strip())
        except ValueError:
            return None
    return None


def render_funeral_md(row, service_date, service_time) -> str:
    venue = row.fields.get("venue")
    wake = row.fields.get("wake_venue")
    provider = row.fields.get("provider")
    lines = ["# Funeral details", "",
             f"- **Date:** {service_date.strftime('%A %d %B %Y')}"]
    if service_time:
        lines.append(f"- **Time:** {service_time.strftime('%H:%M')}")
    if venue:
        lines.append(f"- **Venue:** {venue}")
    if wake:
        lines.append(f"- **Wake afterwards:** {wake}")
    if provider:
        lines.append(f"- **Funeral director:** {provider}")
    lines += ["",
              "A calendar invitation (`funeral-details.ics`) accompanies "
              "this note — import it into any calendar app.", "",
              "> Generated from the funeral arrangements only — safe to "
              "share with the wider circle.", ""]
    return "\n".join(lines)


def render_funeral_ics(row, service_date, service_time) -> str:
    """One VEVENT. DTSTAMP is pinned to the service date so reruns are
    byte-identical; the stable UID makes a re-shared file update, not
    duplicate. Times are floating local time (a UK funeral)."""
    dtstamp = service_date.strftime("%Y%m%dT000000Z")
    if service_time:
        start = datetime.combine(service_date, service_time)
        dt_lines = [
            f"DTSTART:{start.strftime('%Y%m%dT%H%M%S')}",
            f"DTEND:{(start + timedelta(hours=1)).strftime('%Y%m%dT%H%M%S')}",
        ]
    else:
        dt_lines = [
            f"DTSTART;VALUE=DATE:{service_date.strftime('%Y%m%d')}",
            f"DTEND;VALUE=DATE:{(service_date + timedelta(days=1)).strftime('%Y%m%d')}",
        ]
    description = []
    if row.fields.get("wake_venue"):
        description.append(f"Wake afterwards: {row.fields['wake_venue']}")
    if row.fields.get("provider"):
        description.append(f"Funeral director: {row.fields['provider']}")
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//crisis-plan//generator//EN",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "BEGIN:VEVENT",
        f"UID:funeral/{row.external_ref}/service",
        f"DTSTAMP:{dtstamp}",
        *dt_lines,
        "SUMMARY:Funeral service",
    ]
    if row.fields.get("venue"):
        lines.append(f"LOCATION:{_escape(row.fields['venue'])}")
    if description:
        lines.append(f"DESCRIPTION:{_escape(chr(10).join(description))}")
    lines += ["END:VEVENT", "END:VCALENDAR"]
    return CRLF.join(_fold(line) for line in lines) + CRLF


def write_funeral_artifact(model, out_dir):
    """Write funeral-details.md + .ics when a dated funeral row exists.

    Returns ({name: Path}, detail): empty dict when there is nothing to
    share yet, with the reason in detail (never silent)."""
    def drop_stale(reason):
        # Regenerating in place: a previously-written artifact whose
        # source rows vanished must not linger as if still current.
        for name in ("funeral-details.md", "funeral-details.ics"):
            stale = Path(out_dir) / name
            if stale.exists():
                stale.unlink()
        return {}, reason

    rows = model.modules.get("funeral", [])
    if not rows:
        return drop_stale("no funeral tab data — nothing to share")
    row = next((r for r in rows
                if _as_date(r.fields.get("service_date"))), None)
    if row is None:
        return drop_stale("no service_date yet — fill it in and re-run "
                          "once the funeral is arranged")
    service_date = _as_date(row.fields.get("service_date"))
    service_time = _as_time(row.fields.get("service_time"))
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    md_path = out_dir / "funeral-details.md"
    ics_path = out_dir / "funeral-details.ics"
    md_path.write_text(
        render_funeral_md(row, service_date, service_time),
        encoding="utf-8")
    ics_path.write_bytes(
        render_funeral_ics(row, service_date, service_time).encode("utf-8"))
    return ({"funeral-details": md_path, "funeral-ics": ics_path},
            "funeral-details.md + funeral-details.ics — share with the "
            "wider circle when ready")


# --- file writing ------------------------------------------------------------

def write_outputs(tasks, notes, source, trigger: date, out_dir,
                  augment=None) -> dict:
    """Write checklist.md + plan.ics into out_dir; return {name: Path}.

    ``augment`` (optional): row planKey -> cached lookup lines, merged
    into the markdown checklists only — never the .ics (SPEC §13)."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    # Harvest the human's ticks BEFORE anything is overwritten/removed.
    checked = read_checked(out_dir)
    md_path = out_dir / "checklist.md"
    ics_path = out_dir / "plan.ics"
    md_path.write_text(render_markdown(tasks, notes, source, trigger,
                                       augment=augment, checked=checked),
                       encoding="utf-8")
    # .ics carries its own CRLF line endings — write bytes, no translation.
    ics_path.write_bytes(render_ics(tasks, trigger).encode("utf-8"))
    written = {"checklist": md_path, "ics": ics_path}
    # Personal checklists when the work is shared out (SPEC §14, AC-R2).
    # The out dir is regenerated in place, so first drop stale personal
    # lists (e.g. a person since removed from the household tab).
    for stale in out_dir.glob("checklist-*.md"):
        stale.unlink()
    assignees = sorted({t.assignee for t in tasks if t.assignee})
    if len(assignees) > 1:
        for person in assignees:
            path = out_dir / f"checklist-{slugify(person)}.md"
            path.write_text(
                render_person_markdown(tasks, source, trigger, person,
                                       augment=augment, checked=checked),
                encoding="utf-8")
            written[f"checklist:{person}"] = path
    return written


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Render the crisis plan to markdown + .ics.")
    ap.add_argument("workbook", help="path to the .xlsx (sample only in dev)")
    ap.add_argument("--date", required=True, metavar="YYYY-MM-DD",
                    help="trigger (death) date")
    ap.add_argument("--out", required=True, metavar="DIR",
                    help="output directory for checklist.md and plan.ics")
    args = ap.parse_args()

    trigger = date.fromisoformat(args.date)
    model = read_workbook(args.workbook)
    tasks, notes = derive_tasks(model, trigger)
    written = write_outputs(tasks, notes, model.source, trigger, args.out)
    print(f"{len(tasks)} tasks, {len(notes)} skip notes")
    for name, path in written.items():
        print(f"  {name}: {path}")


if __name__ == "__main__":
    main()
