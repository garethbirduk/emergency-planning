#!/usr/bin/env python3
"""Google Calendar reconcile engine (build steps 3.3–3.4).

Upserts the plan's external surface into a calendar, keyed by
``extendedProperties.private.planKey`` (SPEC §5):

- **create** missing events · **update** changed plan fields · **skip**
  unchanged · **flag orphans** (tagged events no longer in the plan get
  a visible ⚠ prefix — kept, never deleted; unflagged if the row
  returns).
- Only events carrying our ``crisisplan=1`` tag are ever fetched or
  touched — untagged items are invisible to us by construction (AC-I5).
- **Status is human-owned** (AC-I6): a leading ``✓`` on the summary
  (the mark-done convention, see GOOGLE-SETUP.md) survives updates, and
  we never send ``colorId``/reminders/attendees, so recolouring is safe.
- Notifications are off on every insert/patch (``sendUpdates=none``,
  AC-M2/M3); sharing happens once at setup, never here.
- Events are the same external surface as the .ics / google-preview:
  titles and descriptions come from the same code paths, so what was
  previewed is what is published (AC-P4).
"""

from datetime import date, timedelta

from outputs import _external_description

DONE_PREFIX = "✓"
ORPHAN_PREFIX = "⚠ [not in plan] "


def desired_events(tasks, run_date: date | None = None) -> dict:
    """planKey -> event body for every externally-visible task.

    ``run_date`` (test/real runs) also schedules the self-canary: an
    'annual review' reminder one year out (SPEC §8), same key every
    run so it reschedules rather than duplicates."""
    desired = {}
    for t in tasks:
        if t.visibility == "none":
            continue
        desired[t.external_key] = {
            "summary": t.external_title,
            "description": _external_description(t),
            "start": {"date": t.due.isoformat()},
            "end": {"date": (t.due + timedelta(days=1)).isoformat()},
            "transparency": "transparent",
            "extendedProperties": {
                "private": {"crisisplan": "1", "planKey": t.external_key}},
        }
    if run_date is not None:
        due = run_date + timedelta(days=365)
        desired["meta/annual-review"] = {
            "summary": "Crisis plan — annual review: re-run the tool",
            "description": ("Yearly canary (see tools/README.md): open "
                            "the workbook, update anything stale, and "
                            "re-run the generator. If this event is the "
                            "only reminder left, the tool has rotted — "
                            "fix it while it's cheap.\n"
                            "Plan key: meta/annual-review"),
            "start": {"date": due.isoformat()},
            "end": {"date": (due + timedelta(days=1)).isoformat()},
            "transparency": "transparent",
            "extendedProperties": {
                "private": {"crisisplan": "1",
                            "planKey": "meta/annual-review"}},
        }
    return desired


def _split_summary(summary: str):
    """-> (done_mark, bare_summary): peel the human's ✓ and our ⚠ flag
    off to compare plan fields underneath."""
    s = summary or ""
    done = ""
    if s.startswith(DONE_PREFIX):
        done = DONE_PREFIX + " "
        s = s[len(DONE_PREFIX):].lstrip()
    if s.startswith(ORPHAN_PREFIX):
        s = s[len(ORPHAN_PREFIX):]
    return done, s


def _needs_update(event: dict, body: dict) -> bool:
    _, bare = _split_summary(event.get("summary"))
    flagged = event.get("extendedProperties", {}).get("private", {}) \
        .get("orphaned") == "true"
    return (flagged
            or bare != body["summary"]
            or (event.get("description") or "") != body["description"]
            or event.get("start", {}).get("date")
            != body["start"]["date"]
            or event.get("end", {}).get("date") != body["end"]["date"])


def reconcile(client, cal_id: str, desired: dict) -> dict:
    """Apply the reconcile rules; return counts for the health line."""
    existing = client.list_tagged_events(cal_id)
    counts = {"created": 0, "updated": 0, "skipped": 0, "orphaned": 0}
    for key in sorted(desired):
        body = desired[key]
        event = existing.pop(key, None)
        if event is None:
            client.insert_event(cal_id, body)
            counts["created"] += 1
        elif not _needs_update(event, body):
            counts["skipped"] += 1
        else:
            done, _ = _split_summary(event.get("summary"))
            client.patch_event(cal_id, event["id"], {
                "summary": done + body["summary"],
                "description": body["description"],
                "start": body["start"],
                "end": body["end"],
                # Clear any orphan flag (None deletes the property).
                "extendedProperties": {"private": {"orphaned": None}},
            })
            counts["updated"] += 1
    # Tagged events no longer in the plan: flag visibly, keep forever.
    for key, event in sorted(existing.items()):
        private = event.get("extendedProperties", {}).get("private", {})
        if private.get("orphaned") == "true":
            counts["skipped"] += 1     # already flagged — leave alone
            continue
        done, bare = _split_summary(event.get("summary"))
        client.patch_event(cal_id, event["id"], {
            "summary": done + ORPHAN_PREFIX + bare,
            "extendedProperties": {"private": {"orphaned": "true"}},
        })
        counts["orphaned"] += 1
    return counts


def run_google_stage(cfg, mode: str, fresh: bool, tasks,
                     run_date: date) -> str:
    """The whole test/real Google stage; returns the health detail.
    Raises GcalError (caller degrades gracefully — AC-D3/M3)."""
    from gcal import CalendarClient, GcalError, get_transport

    client = CalendarClient(get_transport(cfg))
    if fresh:
        cal_id = client.create_calendar(
            f"⚠ TEST crisis-plan {run_date.isoformat()} — DELETE ME")
        target = f"fresh disposable calendar ({cal_id})"
    else:
        attr = "test_calendar_id" if mode == "test" else "real_calendar_id"
        cal_id = getattr(cfg, attr, None) if cfg else None
        if not cal_id:
            raise GcalError(f"{attr} not set — run python tools/gsetup.py "
                            "(see tools/GOOGLE-SETUP.md)")
        target = f"{mode} calendar"
    counts = reconcile(client, cal_id, desired_events(tasks, run_date))
    return (f"{target}: {counts['created']} created, "
            f"{counts['updated']} updated, {counts['skipped']} skipped, "
            f"{counts['orphaned']} orphan-flagged")
