#!/usr/bin/env python3
"""Google config loader (build step 3.1 — AC-M6, AC-P2).

Calendar IDs and credential/token *paths* live in one small JSON file
**outside** the repo; nothing sensitive is ever tracked. This module is
pure stdlib — importing it can never make a network call, so `local`
mode stays offline by construction (AC-M1/D2).

Config file location, in order:
  1. the ``CRISISPLAN_GOOGLE_CONFIG`` environment variable, if set;
  2. ``%APPDATA%\\crisis-plan\\google.json`` on Windows,
     ``~/.config/crisis-plan/google.json`` elsewhere.

Recognised keys (all optional until Phase 3 needs them; unknown keys are
ignored, same tolerance as the workbook reader):
  test_calendar_id   – persistent test calendar (mode ``test``)
  real_calendar_id   – the shared family calendar (mode ``real``)
  credentials_file   – path to the OAuth client credentials JSON
  token_file         – path to the stored user token

A template with dummy values is committed as
``tools/google.config.example.json`` — copy it to the location above and
fill it in during the one-time setup (step 3.2). Paths inside the repo
working tree are refused, mirroring the CLI guards (AC-P3 philosophy).

Usage:  python tools/gconfig.py     (report what would be loaded)
"""

import json
import os
import sys
from dataclasses import dataclass
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
ENV_VAR = "CRISISPLAN_GOOGLE_CONFIG"


@dataclass
class GoogleConfig:
    path: Path
    test_calendar_id: str | None = None
    real_calendar_id: str | None = None
    credentials_file: Path | None = None
    token_file: Path | None = None


def config_path(override=None) -> Path:
    """Explicit override (e.g. generate.py's --google-config, one config per
    household) beats the env var beats the default location."""
    if override:
        return Path(override)
    env = os.environ.get(ENV_VAR)
    if env:
        return Path(env)
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA",
                                   Path.home() / "AppData" / "Roaming"))
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME",
                                   Path.home() / ".config"))
    return base / "crisis-plan" / "google.json"


def load_config(override=None):
    """-> (GoogleConfig | None, detail). Never raises: a missing or
    broken config is a health-summary line, not a crash — Google is the
    fragile tier and must not stop local artifacts (AC-D3)."""
    path = config_path(override)
    if path.resolve().is_relative_to(REPO_ROOT):
        return None, (f"REFUSED: config {path} is inside the repo "
                      "working tree — keep it outside (see PRIVACY.md)")
    if not path.exists():
        return None, (f"no config at {path} (fine for local mode; "
                      "created in the one-time setup, step 3.2)")
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError("not a JSON object")
    except (OSError, ValueError) as e:
        return None, f"config {path} unreadable ({e}) — ignored"

    def _path_field(key):
        value = data.get(key)
        return Path(str(value)).expanduser() if value else None

    cfg = GoogleConfig(
        path=path,
        test_calendar_id=data.get("test_calendar_id") or None,
        real_calendar_id=data.get("real_calendar_id") or None,
        credentials_file=_path_field("credentials_file"),
        token_file=_path_field("token_file"),
    )
    for label in ("credentials_file", "token_file"):
        p = getattr(cfg, label)
        if p is not None and p.resolve().is_relative_to(REPO_ROOT):
            return None, (f"REFUSED: {label} {p} is inside the repo "
                          "working tree — keep credentials outside "
                          "(see PRIVACY.md)")
    ids_set = [key for key in ("test_calendar_id", "real_calendar_id")
               if getattr(cfg, key)]
    if ids_set:
        detail = f"loaded {path.name} ({', '.join(ids_set)} set)"
    else:
        detail = f"loaded {path.name} (no calendar ids set yet)"
    return cfg, detail


def main() -> None:
    cfg, detail = load_config()
    print(f"Config location: {config_path()}")
    print(f"Status: {detail}")
    if cfg is None:
        sys.exit(1 if detail.startswith("REFUSED") else 0)
    for field in ("test_calendar_id", "real_calendar_id",
                  "credentials_file", "token_file"):
        print(f"  {field}: {getattr(cfg, field) or '(not set)'}")


if __name__ == "__main__":
    main()
