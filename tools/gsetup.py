#!/usr/bin/env python3
"""One-time Google Calendar setup (build step 3.2 — AC-M5, AC-M6).

The deliberate, separate act that normal runs must never perform: sign
in once, create the two calendars, share the real one with the family,
and write the config. Everything sensitive lands OUTSIDE the repo.

- **Re-runnable / idempotent:** an existing config is kept; calendars
  whose IDs still resolve are reused, not duplicated; shares are
  re-asserted (upsert by email). Run it again any time to repair.
- **Prerequisites** (tools/GOOGLE-SETUP.md walks through them): an OAuth
  desktop-client JSON downloaded from Google Cloud Console, saved
  outside the repo. The browser consent runs from here.
- The real calendar is shared as calendar *membership* (never per-event
  guests), so re-running the plan can never spam anyone (SPEC §6).

Usage:
  python tools/gsetup.py [--emails a@x,b@y] [--household NAME] [--dry-run]
"""

import argparse
import json
import sys
from pathlib import Path

import gconfig
from gcal import (CalendarClient, GcalError, get_transport,
                  interactive_authorize)

REPO_ROOT = Path(__file__).resolve().parents[1]


def _read_line(prompt: str) -> str:
    print(prompt, end="", flush=True)
    return input().replace("﻿", "").strip()


def _existing_raw(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def main() -> None:
    ap = argparse.ArgumentParser(
        prog="gsetup.py",
        description="One-time Google Calendar setup: sign in, create + "
                    "share the calendars, write the config.")
    ap.add_argument("--emails", metavar="A@X,B@Y",
                    help="family members to share the real calendar with "
                         "(comma-separated; prompted for if omitted)")
    ap.add_argument("--household", metavar="NAME",
                    help="label for the real calendar (default: "
                         "'Crisis plan')")
    ap.add_argument("--dry-run", action="store_true",
                    help="print what would be done; change nothing, "
                         "touch no network")
    args = ap.parse_args()

    cfg_path = gconfig.config_path()
    if cfg_path.resolve().is_relative_to(REPO_ROOT):
        sys.exit(f"REFUSED: config path {cfg_path} is inside the repo "
                 "working tree — keep it outside (see PRIVACY.md)")
    raw = _existing_raw(cfg_path)
    secrets_dir = cfg_path.parent
    credentials_file = Path(raw.get("credentials_file")
                            or secrets_dir / "oauth-client.json")
    token_file = Path(raw.get("token_file") or secrets_dir / "token.json")
    household = args.household or "Crisis plan"
    real_name = household if "crisis" in household.lower() \
        else f"Crisis plan — {household}"

    print("One-time Google Calendar setup")
    print(f"  config file:      {cfg_path}"
          + (" (exists — will be updated)" if raw else " (will be created)"))
    print(f"  oauth client:     {credentials_file}")
    print(f"  token store:      {token_file}")
    print(f"  test calendar:    "
          + (raw.get("test_calendar_id") or "(will be created)"))
    print(f"  real calendar:    "
          + (raw.get("real_calendar_id")
             or f"(will be created as '{real_name}')"))
    if args.dry_run:
        print("\n--dry-run: nothing was changed, nothing was sent.")
        return

    emails = args.emails
    if emails is None:
        emails = _read_line(
            "Family emails to share the real calendar with "
            "(comma-separated, blank to skip): ")
    emails = [e.strip() for e in (emails or "").split(",") if e.strip()]

    # Sign in (once). The fake test transport needs no credentials.
    import os
    if not os.environ.get("CRISISPLAN_GCAL_FAKE"):
        if not credentials_file.exists():
            sys.exit(f"No OAuth client JSON at {credentials_file}.\n"
                     "Follow tools/GOOGLE-SETUP.md step 1 to create and "
                     "download it, then re-run this setup.")
        if not token_file.exists():
            interactive_authorize(credentials_file, token_file)
            print("Authorised; refresh token stored.")

    cfg = gconfig.GoogleConfig(path=cfg_path,
                               credentials_file=credentials_file,
                               token_file=token_file)
    client = CalendarClient(get_transport(cfg))

    test_id = raw.get("test_calendar_id")
    if test_id and client.calendar_exists(test_id):
        print(f"Test calendar kept: {test_id}")
    else:
        if test_id:
            print(f"Test calendar {test_id} no longer exists — "
                  "creating a new one.")
        test_id = client.create_calendar(
            "Crisis plan — TEST (private, never shared)")
        print(f"Test calendar created: {test_id}")
    # Repair the 'unsubscribed but not deleted' case: put the calendar
    # back in the owner's visible list (no-op when already there).
    client.ensure_listed(test_id)

    real_id = raw.get("real_calendar_id")
    if real_id and client.calendar_exists(real_id):
        print(f"Real calendar kept: {real_id}")
    else:
        if real_id:
            print(f"Real calendar {real_id} no longer exists — "
                  "creating a new one.")
        real_id = client.create_calendar(real_name)
        print(f"Real calendar created: {real_id}")
    client.ensure_listed(real_id)
    for email in emails:
        try:
            client.share(real_id, email)
            print(f"  shared with {email} (writer)")
        except GcalError as e:
            print(f"  share with {email} FAILED: {e}")

    cfg_path.parent.mkdir(parents=True, exist_ok=True)
    cfg_path.write_text(json.dumps({
        **{k: v for k, v in raw.items()
           if k not in ("test_calendar_id", "real_calendar_id",
                        "credentials_file", "token_file")},
        "test_calendar_id": test_id,
        "real_calendar_id": real_id,
        "credentials_file": str(credentials_file),
        "token_file": str(token_file),
    }, indent=2), encoding="utf-8")
    print(f"\nConfig written: {cfg_path}")
    print("Try it:  python tools\\generate.py -f <workbook> --date "
          "YYYY-MM-DD --mode test")


if __name__ == "__main__":
    main()
