# Google Calendar setup — one-time checklist

Everything here happens **once**, takes ~15 minutes, and every sensitive
file lands **outside** this repo (in `%APPDATA%\crisis-plan\`). Until you
do this, the tool works fully offline — `google-preview.md` already
shows you exactly what a push would publish.

## 1. Create the OAuth client (Google Cloud Console)

1. Go to <https://console.cloud.google.com/> signed in as the account
   that should own the family calendar.
2. Create a project (any name, e.g. `crisis-plan`).
3. **APIs & Services → Library** → enable **Google Calendar API**.
4. **APIs & Services → OAuth consent screen** → External → fill in the
   two required fields. Then **publish the app** (Testing mode expires
   refresh tokens after 7 days — published mode does not; the scary
   "unverified app" warning is fine, it's your own app used only by
   you).
5. **APIs & Services → Credentials → Create credentials → OAuth client
   ID** → application type **Desktop app**.
6. Download the JSON and save it as:
   `%APPDATA%\crisis-plan\oauth-client.json`

## 2. Create the config file

Copy `tools\google.config.example.json` to
`%APPDATA%\crisis-plan\google.json` and set the two file paths (the
calendar IDs can stay as placeholders — step 3 fills them in):

```json
{
  "credentials_file": "C:/Users/<you>/AppData/Roaming/crisis-plan/oauth-client.json",
  "token_file": "C:/Users/<you>/AppData/Roaming/crisis-plan/token.json"
}
```

(To keep the config somewhere else entirely, set the
`CRISISPLAN_GOOGLE_CONFIG` environment variable to its full path — handy
for one-config-per-household later.)

## 3. Run the setup command

```
python tools\gsetup.py --emails alex@example.com,chris@example.com
```

It opens the browser for consent (once), creates a private **test**
calendar and the shared **real** calendar, shares the real one with the
emails you gave (calendar membership — nobody ever gets per-event
invites), and writes both IDs into the config. Re-running it is safe:
existing calendars are kept and repaired, never duplicated. `--dry-run`
shows the plan without touching anything.

## 4. Verify

```
python tools\generate.py -f <workbook> --date YYYY-MM-DD --mode test
```

pushes into the private test calendar only. Check it in Google
Calendar, then use `--mode real` when you're happy. The health summary
reports created/updated/skipped counts every run, and a Google failure
never blocks the local outputs.

## If a calendar "disappears"

In Google Calendar's settings an owned calendar has two removal buttons:
**Unsubscribe** (hides it from your list — the calendar and its events
live on) and **Delete** (destroys it). If one seems to have vanished,
just re-run `python tools\gsetup.py` — it detects that the calendar
still exists and puts it back in your list. If you genuinely deleted
it, the same re-run creates a fresh one and updates the config.

**Deletions propagate slowly** (observed live: several minutes). If you
delete a calendar and immediately re-run setup, Google's API may still
report it as existing — setup will then say "kept" for a calendar
that's actually dying. Wait a few minutes and re-run; once the API
returns 404 the repair creates a fresh calendar. Re-runs are always
safe, so when in doubt, run it again.

## Marking tasks done (the convention)

Calendar events have no checkbox. The convention: **prefix the event
title with `✓`** when done (or recolour it). Re-runs preserve the `✓`
and never touch colours — the plan updates dates/wording underneath
your mark. The printed / markdown checklist remains the authoritative
done-list.
