#!/usr/bin/env python3
"""AI augmentation tool (build step 2.3, SPEC §13).

An entirely **opt-in** extra: for workbook rows the household has marked
``ai_visible? = Y``, look up practical bereavement extras (URLs, phone
lines) for the row's institution/service and cache them in a sidecar file.
The renderers then merge the cached lines into the *local* checklist as
extra detail. Everything works identically without the cache.

Privacy design (AC-V4, AC-V5):

- **Minimised queries** — a query is the institution/service name plus a
  fixed task-context phrase, nothing else. Owner, account label, location
  and note fields are never read into a query. Rows without
  ``ai_visible? = Y`` contribute nothing at all.
- **Consent gate** — the exact query list is printed and the run stops for
  explicit confirmation ("yes") before any lookup happens. The gate output
  identifies rows only as ``<tab> #<id>``.
- **Pluggable lookup, offline first** — providers:
    ``stub``   (default) offline placeholder lines; proves the plumbing
               end-to-end with no network and no keys.
    ``manual`` the documented paste flow: the tool shows you each query,
               YOU search the web in your own browser, paste the useful
               lines back in (blank line to finish, nothing to skip).
  A live search provider can slot in later behind the same gate.
- **Sidecar cache, untracked** — results are cached keyed by the row's
  planKey, *outside* git: beside the real workbook
  (``<workbook>.augment.json``) or, for a workbook inside the repo (the
  sample), under gitignored ``private/augment/``. The workbook itself is
  never written (AC-P1).
- **Local merge only** — cached lines appear in the markdown checklists
  only, never in the ``.ics``: the calendar is an external surface and
  institution-specific lines would defeat a ``generic``/``none`` row's
  redaction.

Usage:
  python tools/augment.py -f <workbook.xlsx>                # gate + stub
  python tools/augment.py -f <workbook.xlsx> --provider manual
  python tools/augment.py -f <workbook.xlsx> --list         # preview only
"""

import argparse
import json
import sys
from datetime import date
from pathlib import Path

from reader import read_workbook

REPO_ROOT = Path(__file__).resolve().parents[1]

# The single column per tab whose value may appear in a query. Tabs not
# listed here have no institution/service concept — their rows are never
# queried, even with ai_visible? = Y.
INSTITUTION_COLUMN = {
    "banks": "bank",
    "building_societies": "society",
    "cash_isas": "provider",
    "shares_isas": "provider",
    "share_portfolios": "registrar_or_platform",
    "pensions": "provider",
    "digital": "service",
    "funeral": "provider",
    "professionals": "name",
    "life_insurance": "provider",
}

# Fixed task-context phrase appended to the institution name — generic on
# purpose; it must never need row detail to be useful.
QUERY_CONTEXT = {
    "banks": "bereavement team contact and process when an account "
             "holder dies (UK)",
    "building_societies": "bereavement team contact and process when an "
                          "account holder dies (UK)",
    "cash_isas": "bereavement process and inherited ISA (APS) allowance (UK)",
    "shares_isas": "bereavement process and inherited ISA (APS) "
                   "allowance (UK)",
    "share_portfolios": "deceased shareholder services and bereavement "
                        "process (UK)",
    "pensions": "bereavement contact and death benefits process (UK)",
    "digital": "closing or memorialising an account when the owner dies",
    "funeral": "funeral director contact details and reviews (UK)",
    "professionals": "firm contact details (UK)",
    "life_insurance": "bereavement claims process and contact (UK)",
}


def default_cache_path(workbook) -> Path:
    """Untracked sidecar location: beside a real (outside-repo) workbook,
    or under gitignored private/ for a workbook inside the repo."""
    wb = Path(workbook).resolve()
    if wb.is_relative_to(REPO_ROOT):
        return REPO_ROOT / "private" / "augment" / f"{wb.stem}.augment.json"
    return wb.with_suffix(".augment.json")


def guard_paths(source: Path, cache: Path) -> None:
    """Same philosophy as generate.py's guard (AC-P3): inside the repo tree
    only the committed sample may be read and only private/ written."""
    src = source.resolve()
    if src.is_relative_to(REPO_ROOT) and \
            not src.is_relative_to(REPO_ROOT / "samples"):
        sys.exit(
            f"REFUSED: -f {src}\n"
            f"is inside the repo working tree ({REPO_ROOT}).\n"
            "The real spreadsheet must live OUTSIDE the repo (see "
            "PRIVACY.md).")
    dst = cache.resolve()
    if dst.is_relative_to(REPO_ROOT) and \
            not dst.is_relative_to(REPO_ROOT / "private"):
        sys.exit(
            f"REFUSED: --cache {dst}\n"
            f"is inside the repo working tree ({REPO_ROOT}).\n"
            "The cache may only go to the gitignored private/ area or a "
            "location outside the repo (beside the real workbook).")


def load_cache(path) -> dict:
    """Read a sidecar cache -> {plan_key: [detail lines]}. Absent file ->
    {}; a corrupt file warns and is treated as absent (the plan must never
    depend on this fragile extra)."""
    path = Path(path)
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        entries = data.get("entries", {})
        return {key: [str(line) for line in entry.get("lines", [])]
                for key, entry in entries.items()
                if entry.get("lines")}
    except (json.JSONDecodeError, AttributeError, TypeError) as e:
        print(f"WARNING: augment cache {path} unreadable ({e}) — ignored",
              file=sys.stderr)
        return {}


def extract_queries(model):
    """Minimised queries from consenting rows only (AC-V4).

    Returns (queries, notes): queries as a list of dicts with plan_key,
    ref ("<tab> #<id>"), institution, and the exact query string; notes
    for consenting rows nothing can be asked about."""
    queries, notes = [], []
    for tab, rows in model.modules.items():
        for row in rows:
            if not row.ai_visible:
                continue
            ref = f"{tab} #{row.external_ref}"
            column = INSTITUTION_COLUMN.get(tab)
            if column is None:
                notes.append(f"{ref}: tab has no institution/service "
                             "column — nothing to look up")
                continue
            name = row.fields.get(column)
            if not name:
                notes.append(f"{ref}: '{column}' cell is empty — "
                             "nothing to look up")
                continue
            queries.append({
                "plan_key": row.plan_key,
                "ref": ref,
                "institution": str(name),
                "query": f"{name} — {QUERY_CONTEXT[tab]}",
            })
    return queries, notes


def print_gate(queries) -> None:
    """The consent gate's query listing: exact queries, rows named only by
    tab + id — no owners, account labels, or locations (AC-V4)."""
    unique = {}
    for q in queries:
        unique.setdefault(q["query"], []).append(q["ref"])
    print("Consent gate — these EXACT queries would be looked up, "
          "nothing more:")
    print()
    for i, (query, refs) in enumerate(unique.items(), start=1):
        print(f'  {i}. "{query}"')
        print(f"     for row(s): {', '.join(refs)}")
    print()
    print("Only the institution/service name and the fixed phrase above "
          "are sent.")
    print("Owners, account labels, locations and notes are never included; "
          "rows without ai_visible? = Y are not used at all.")


def stub_lookup(query: str, institution: str) -> list:
    """Offline placeholder provider — no network, no keys. Proves the
    gate -> cache -> renderer plumbing; replace via --provider manual."""
    return [
        f"[ai-stub] Placeholder for '{institution}' — no lookup was made "
        "(offline stub provider).",
        "[ai-stub] Re-run tools/augment.py with --provider manual to add "
        "real bereavement links.",
    ]


def manual_lookup(query: str, institution: str) -> list:
    """The documented manual paste flow: you search, you paste."""
    print(f"\nSearch the web yourself (your own browser) for:")
    print(f'  "{query}"')
    print("Paste useful lines (bereavement URL, phone line…), one per "
          "line.")
    print("Finish with a blank line; paste nothing to skip this one.")
    lines = []
    while True:
        try:
            line = _read_line()
        except EOFError:
            break
        if not line.strip():
            break
        lines.append(line.strip())
    return lines


def _read_line() -> str:
    """input(), minus the UTF-8 BOM a PowerShell pipe prefixes to stdin."""
    return input().replace("﻿", "")


PROVIDERS = {"stub": stub_lookup, "manual": manual_lookup}


def write_cache(path: Path, workbook: Path, results: dict) -> None:
    """Merge results into the sidecar cache (existing entries for other
    rows are kept). The workbook is never touched."""
    path = Path(path)
    existing = {}
    if path.exists():
        try:
            existing = json.loads(
                path.read_text(encoding="utf-8")).get("entries", {})
        except (json.JSONDecodeError, AttributeError):
            existing = {}  # corrupt — rebuild from scratch
    existing.update(results)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps({
        "version": 1,
        "generated": date.today().isoformat(),
        "workbook": Path(workbook).name,
        "note": "Sidecar cache written by tools/augment.py — untracked, "
                "safe to delete; the plan renders fine without it.",
        "entries": existing,
    }, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def main() -> None:
    ap = argparse.ArgumentParser(
        prog="augment.py",
        description="Opt-in lookup of practical extras for ai_visible? = Y "
                    "rows, cached in an untracked sidecar file.")
    ap.add_argument("-f", "--file", required=True, metavar="XLSX",
                    help="input spreadsheet (real one lives outside the "
                         "repo)")
    ap.add_argument("--provider", choices=sorted(PROVIDERS),
                    default="stub",
                    help="lookup provider: offline stub (default) or the "
                         "manual search-and-paste flow")
    ap.add_argument("--cache", metavar="PATH",
                    help="sidecar cache file (default: beside the real "
                         "workbook, or private/augment/ for the sample)")
    ap.add_argument("--list", action="store_true",
                    help="print the consent-gate query list and exit — "
                         "no confirmation, no lookup, no cache write")
    args = ap.parse_args()

    source = Path(args.file)
    cache_path = Path(args.cache) if args.cache \
        else default_cache_path(source)
    guard_paths(source, cache_path)

    model = read_workbook(source)  # read-only; never written (AC-P1)
    queries, notes = extract_queries(model)
    for note in notes:
        print(f"note: {note}")
    if not queries:
        print("No rows with ai_visible? = Y have anything to look up — "
              "nothing to do.")
        return

    print_gate(queries)
    if args.list:
        return
    print()
    print("Proceed with the lookup? Type 'yes' to continue: ", end="",
          flush=True)
    answer = _read_line()
    if answer.strip().lower() != "yes":
        sys.exit("Aborted — nothing was looked up; the cache was not "
                 "touched.")

    lookup = PROVIDERS[args.provider]
    lines_by_query = {}   # one lookup per unique query
    results, skipped = {}, []
    for q in queries:
        if q["query"] not in lines_by_query:
            lines_by_query[q["query"]] = lookup(q["query"],
                                                q["institution"])
        lines = lines_by_query[q["query"]]
        if lines:
            results[q["plan_key"]] = {"query": q["query"], "lines": lines}
        else:
            skipped.append(q["ref"])

    if not results:
        sys.exit("No results to cache (every query was skipped); the "
                 "cache was not touched.")
    write_cache(cache_path, source, results)
    print(f"\nCached {len(results)} row(s) -> {cache_path}")
    for ref in skipped:
        print(f"  skipped (no lines): {ref}")
    print("Re-run the plan (tools/generate.py) to merge the extras into the "
          "checklist.")


if __name__ == "__main__":
    main()
