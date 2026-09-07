#!/usr/bin/env python3
"""Crisis-plan CLI (build step 1.5).

The single entry point (SPEC §10): workbook + trigger date in, dated
snapshot out, with a per-stage health summary (AC-D4) so a future run tells
you what has rotted while it's cheap to fix.

    python tools/generate.py -f <person>.xlsx --date YYYY-MM-DD
                           [--mode {local,test,real}] [--fresh] [--out DIR]

- **Snapshot (AC-D1):** every run writes ``checklist.md`` + ``plan.ics`` +
  ``run-info.txt`` to ``private/runs/<run-date>/`` (gitignored), or ``--out``.
- **Repo guard (AC-P3):** refuses ``-f``/``--out`` paths inside the repo
  working tree. Deliberate exceptions, because they are safe by design:
  ``-f`` under ``samples/`` (committed dummy data) and ``--out`` under
  ``private/`` (gitignored — the default snapshot area).
- **Local first (AC-D3):** local artifacts are written before the Google
  stage runs. ``local`` mode touches no network (AC-D2/M1) — the
  network-capable module (gcal, via reconcile) is imported lazily and
  only for ``test``/``real``.
- A stage failure is reported in the summary, dependent stages are skipped,
  and the exit code is non-zero — the summary always prints. The Google
  stage is the exception (SPEC §9): its failures are WARN, the local
  artifacts stand, and the run still succeeds.
"""

import argparse
import sys
from datetime import date
from pathlib import Path

from augment import default_cache_path, load_cache
from gconfig import load_config
from reader import read_workbook
from tasks import derive_tasks
from outputs import (write_funeral_artifact, write_google_preview,
                     write_outputs)
from prep import write_preparation

REPO_ROOT = Path(__file__).resolve().parents[1]

OK, FAIL, SKIP, WARN = "OK", "FAIL", "SKIP", "WARN"


class Health:
    """Per-stage outcomes for the end-of-run summary (AC-D4)."""

    def __init__(self):
        self.stages = []

    def record(self, stage: str, status: str, detail: str = ""):
        self.stages.append((stage, status, detail))

    @property
    def failed(self) -> bool:
        return any(status == FAIL for _, status, _ in self.stages)

    def report(self) -> str:
        lines = ["", "Health summary"]
        for stage, status, detail in self.stages:
            dots = "." * max(2, 22 - len(stage))
            lines.append(f"  {stage} {dots} {status}"
                         + (f" — {detail}" if detail else ""))
        lines.append(f"  overall {'.' * 15} "
                     + ("FAIL" if self.failed else "OK"))
        return "\n".join(lines)


def guard_paths(source: Path, out_dir: Path) -> None:
    """Refuse repo-tree paths (AC-P3) so real data / outputs derived from
    it can never land in the tracked tree by accident."""
    src = source.resolve()
    if src.is_relative_to(REPO_ROOT) and \
            not src.is_relative_to(REPO_ROOT / "samples"):
        sys.exit(
            f"REFUSED: -f {src}\n"
            f"is inside the repo working tree ({REPO_ROOT}).\n"
            "The real spreadsheet must live OUTSIDE the repo (see "
            "PRIVACY.md); only the committed dummy data under samples/ is "
            "allowed from inside it.")
    out = out_dir.resolve()
    if out.is_relative_to(REPO_ROOT) and \
            not out.is_relative_to(REPO_ROOT / "private"):
        sys.exit(
            f"REFUSED: --out {out}\n"
            f"is inside the repo working tree ({REPO_ROOT}).\n"
            "Outputs may only go to the gitignored private/ area or a "
            "directory outside the repo.")


def write_run_info(out_dir: Path, args, trigger: date, tasks, notes) -> None:
    """A snapshot is self-describing: what ran, from what, when."""
    (out_dir / "run-info.txt").write_text(
        "Crisis-plan run\n"
        f"  run date:   {date.today().isoformat()}\n"
        f"  source:     {Path(args.file).name}\n"
        f"  trigger:    {trigger.isoformat()}\n"
        f"  mode:       {args.mode}\n"
        f"  tasks:      {len(tasks)}\n"
        f"  skip notes: {len(notes)}\n",
        encoding="utf-8")


def main() -> None:
    ap = argparse.ArgumentParser(
        prog="generate.py",
        description="Generate a dated crisis plan (checklist + calendar) "
                    "from a person's data spreadsheet.")
    ap.add_argument("-f", "--file", required=True, metavar="XLSX",
                    help="input spreadsheet (real one lives outside the repo)")
    ap.add_argument("--date", required=True, metavar="YYYY-MM-DD",
                    help="trigger (death) date all task dates count from")
    ap.add_argument("--mode", choices=["local", "test", "real"],
                    default="local",
                    help="local = offline, no Google (default until Phase 3)")
    ap.add_argument("--fresh", action="store_true",
                    help="use a disposable test calendar (Phase 3; ignored "
                         "in local mode)")
    ap.add_argument("--out", metavar="DIR",
                    help="output directory (default: private/runs/<run-date>)")
    ap.add_argument("--augment-cache", metavar="PATH",
                    help="sidecar cache from tools/augment.py (default: "
                         "its standard location for this workbook)")
    ap.add_argument("--google-config", metavar="PATH",
                    help="google config file for THIS household "
                         "(default: CRISISPLAN_GOOGLE_CONFIG env var, "
                         "else %%APPDATA%%\\crisis-plan\\google.json)")
    args = ap.parse_args()

    try:
        trigger = date.fromisoformat(args.date)
    except ValueError:
        ap.error(f"--date {args.date!r} is not a valid YYYY-MM-DD date")

    if args.out:
        out_dir = Path(args.out)
    else:
        source = Path(args.file).resolve()
        if source.is_relative_to(REPO_ROOT):
            # Sample/dev runs: dated snapshots in the gitignored area.
            out_dir = (REPO_ROOT / "private" / "runs"
                       / date.today().isoformat())
        else:
            # A real workbook's plan lives in the same folder as the
            # workbook (one workbook per folder), same place the wizard
            # links (AC-D1).
            out_dir = source.parent
    guard_paths(Path(args.file), out_dir)

    health = Health()
    print(f"Crisis plan: {Path(args.file).name} → {out_dir}  "
          f"(trigger {trigger}, mode {args.mode})")

    model = tasks = notes = None
    try:
        model = read_workbook(args.file)
        health.record("read workbook", OK,
                      f"{len(model.modules)} module tab(s) with data")
    except Exception as e:
        health.record("read workbook", FAIL, str(e))

    if model is not None:
        try:
            tasks, notes = derive_tasks(model, trigger)
            health.record("derive tasks", OK,
                          f"{len(tasks)} tasks, {len(notes)} skip note(s)")
        except Exception as e:
            health.record("derive tasks", FAIL, str(e))

    # Optional augment sidecar (SPEC §13) — a fragile extra: merged into
    # the markdown when present, and its absence changes nothing.
    augment_data = {}
    cache_path = Path(args.augment_cache) if args.augment_cache \
        else default_cache_path(Path(args.file))
    if cache_path.exists():
        augment_data = load_cache(cache_path)
        if augment_data:
            health.record("augment cache", OK,
                          f"{len(augment_data)} row(s) from "
                          f"{cache_path.name}")
        else:
            health.record("augment cache", SKIP,
                          f"{cache_path.name} empty or unreadable — "
                          "ignored")
    else:
        health.record("augment cache", SKIP,
                      "no cache — tools/augment.py adds lookup extras")

    if tasks is not None:
        try:
            written = write_outputs(tasks, notes, model.source, trigger,
                                    out_dir, augment=augment_data)
            write_run_info(out_dir, args, trigger, tasks, notes)
            health.record("local outputs", OK,
                          ", ".join(p.name for p in written.values())
                          + " + run-info.txt")
        except Exception as e:
            health.record("local outputs", FAIL, str(e))
    else:
        health.record("local outputs", SKIP, "nothing to write")

    # Preparation plan (SPEC §16 first slice): the local-only "what to
    # do before someone dies" markdown — declared rows + noticed gaps.
    if model is not None:
        try:
            _, detail = write_preparation(model, out_dir)
            health.record("preparation", OK, detail)
        except Exception as e:
            health.record("preparation", FAIL, str(e))
    else:
        health.record("preparation", SKIP, "nothing to write")

    # Friends-facing funeral artifact (SPEC §7): a separate share-with-
    # the-wider-circle file pair, from the funeral tab only, once a
    # service date is known.
    if model is not None:
        try:
            extra, detail = write_funeral_artifact(model, out_dir)
            health.record("friends artifact", OK if extra else SKIP,
                          detail)
        except Exception as e:
            health.record("friends artifact", FAIL, str(e))
    else:
        health.record("friends artifact", SKIP, "nothing to write")

    # Google config (AC-M6): IDs/creds live outside the repo. Reported in
    # every mode so rot is visible early; only test/real will need it.
    cfg, cfg_detail = load_config(args.google_config)
    cfg_status = FAIL if cfg_detail.startswith("REFUSED") \
        else (OK if cfg else SKIP)
    health.record("google config", cfg_status, cfg_detail)

    # Publish preview (AC-M7): what WOULD go to Google, written before
    # any network stage exists — the family reads this, not API calls.
    if tasks is not None:
        try:
            if cfg is None:
                calendar_note = ("not configured yet — see "
                                 "tools/README.md (Google setup)")
            else:
                ids = [k for k in ("test_calendar_id", "real_calendar_id")
                       if getattr(cfg, k)]
                calendar_note = (f"config loaded ({', '.join(ids)} set)"
                                 if ids else
                                 "config loaded, no calendar ids yet")
            _, detail = write_google_preview(tasks, trigger, args.mode,
                                             calendar_note, out_dir)
            health.record("google preview", OK, detail)
        except Exception as e:
            health.record("google preview", FAIL, str(e))
    else:
        health.record("google preview", SKIP, "nothing to write")

    if args.mode == "local":
        health.record("google", SKIP, "local mode — no network")
    elif tasks is None:
        health.record("google", SKIP, "nothing to publish")
    else:
        # The fragile tier (SPEC §9): any failure here is reported as a
        # WARN — the local artifacts above are already complete and the
        # run still succeeds (AC-D3).
        try:
            from reconcile import run_google_stage  # lazy: keeps local
            # mode free of any network-capable import (AC-M1)
            detail = run_google_stage(cfg, args.mode, args.fresh, tasks,
                                      date.today())
            health.record("google", OK, detail)
        except Exception as e:
            health.record("google", WARN,
                          f"{e} — local artifacts above are complete")

    print(health.report())
    sys.exit(1 if health.failed else 0)


if __name__ == "__main__":
    main()
