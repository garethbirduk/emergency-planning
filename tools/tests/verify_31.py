"""Regression suite from build step 3.1: Google config scaffolding +
publish preview (AC-M6, AC-M7, AC-P2).

The loader reads calendar IDs / credential paths from outside the repo
and never raises; google-preview.md is written every run and mirrors the
.ics external surface one-to-one.
"""

import json
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SAMPLE = REPO / "samples" / "pat.sample.xlsx"
SCRATCH = Path(tempfile.mkdtemp(prefix="crisisplan-verify31-"))
sys.path.insert(0, str(REPO / "tools"))

failures = []


def check(name, ok, detail=""):
    print(f"  [{'PASS' if ok else 'FAIL'}] {name}" + (f" — {detail}" if detail else ""))
    if not ok:
        failures.append(name)


def run_cli(out_dir, env_config=None, expect_ok=True):
    env = dict(os.environ)
    env.pop("CRISISPLAN_GOOGLE_CONFIG", None)
    if env_config is not None:
        env["CRISISPLAN_GOOGLE_CONFIG"] = str(env_config)
    r = subprocess.run(
        [sys.executable, str(REPO / "tools" / "generate.py"), "-f",
         str(SAMPLE), "--date", "2030-01-15", "--out", str(out_dir)],
        capture_output=True, text=True, env=env)
    if expect_ok and r.returncode != 0:
        print(r.stdout, r.stderr)
        raise SystemExit("CLI failed")
    return r


print("AC-M7: preview written every run, mirrors the .ics exactly")
out = SCRATCH / "run-a"
r = run_cli(out)
preview_path = out / "google-preview.md"
check("google-preview.md written", preview_path.exists())
check("health reports preview + config status",
      "google preview" in r.stdout and "google config" in r.stdout)
preview = preview_path.read_text(encoding="utf-8")
ics = (out / "plan.ics").read_bytes().decode("utf-8").replace("\r\n ", "")
ics_uids = set(re.findall(r"^UID:([^\r\n]+)", ics, re.M))
preview_keys = set(re.findall(r"^  - Plan key: (.+)$", preview, re.M))
check("preview keys == .ics UIDs, one-to-one",
      preview_keys == ics_uids,
      f"{len(preview_keys)} preview vs {len(ics_uids)} ics")
check("counts line matches", f"**Events to publish:** {len(ics_uids)}"
      in preview)

print("AC-M7: preview obeys the visibility policy")
check("generic rows redacted (no Sampletown)",
      "Sampletown" not in preview
      and "Notify building society #1" in preview)
check("full row keeps its real name",
      "Notify Example Bank plc — Current account (bank #1)" in preview)
check("none row only as a withheld reference",
      "Family home" not in preview and "Example Street" not in preview
      and "`property #1` — 2 task(s), local checklist only" in preview)
check("redacted descriptions only point home",
      "Passbook" not in preview and "Password manager" not in preview)

print("Determinism")
out_b = SCRATCH / "run-b"
run_cli(out_b)
check("preview byte-identical across reruns",
      preview_path.read_bytes()
      == (out_b / "google-preview.md").read_bytes())

print("AC-M6: config loader (in-process)")
import gconfig  # noqa: E402

_saved = os.environ.get(gconfig.ENV_VAR)


def set_env(value):
    if value is None:
        os.environ.pop(gconfig.ENV_VAR, None)
    else:
        os.environ[gconfig.ENV_VAR] = str(value)


good = SCRATCH / "google.json"
good.write_text(json.dumps({
    "test_calendar_id": "dummy-test@group.calendar.google.com",
    "credentials_file": str(SCRATCH / "oauth-client.json"),
    "unknown_key": "ignored",
}), encoding="utf-8")
try:
    set_env(good)
    cfg, detail = gconfig.load_config()
    check("valid config loads via env var",
          cfg is not None
          and cfg.test_calendar_id == "dummy-test@group.calendar.google.com"
          and cfg.real_calendar_id is None
          and "test_calendar_id set" in detail, detail)
    check("unknown keys tolerated", cfg is not None)

    set_env(SCRATCH / "does-not-exist.json")
    cfg, detail = gconfig.load_config()
    check("missing config -> None + reason",
          cfg is None and "no config at" in detail)

    bad = SCRATCH / "bad.json"
    bad.write_text("not json", encoding="utf-8")
    set_env(bad)
    cfg, detail = gconfig.load_config()
    check("corrupt config -> None + 'unreadable', no crash",
          cfg is None and "unreadable" in detail)

    inside = REPO / "private" / "sneaky.json"
    set_env(inside)
    cfg, detail = gconfig.load_config()
    check("config path inside the repo refused",
          cfg is None and detail.startswith("REFUSED"))

    creds_inside = SCRATCH / "creds-inside.json"
    creds_inside.write_text(json.dumps(
        {"credentials_file": str(REPO / "tools" / "oauth.json")}),
        encoding="utf-8")
    set_env(creds_inside)
    cfg, detail = gconfig.load_config()
    check("credentials path inside the repo refused",
          cfg is None and detail.startswith("REFUSED")
          and "credentials_file" in detail)
finally:
    set_env(_saved)

print("AC-M6 via the CLI")
r = run_cli(SCRATCH / "run-c", env_config=good)
check("CLI health: config OK with the env-var file",
      re.search(r"google config \.+ OK", r.stdout) is not None
      and "test_calendar_id set" in r.stdout)
r = run_cli(SCRATCH / "run-d", env_config=REPO / "private" / "x.json",
            expect_ok=False)
check("CLI fails loudly on an inside-repo config",
      r.returncode != 0 and "REFUSED" in r.stdout)
check("local artifacts still written despite the config FAIL",
      (SCRATCH / "run-d" / "checklist.md").exists()
      and (SCRATCH / "run-d" / "google-preview.md").exists())

print("AC-P2: committed example is dummy-only")
example = REPO / "tools" / "google.config.example.json"
data = json.loads(example.read_text(encoding="utf-8"))
check("example parses and every value is an obvious placeholder",
      all("EXAMPLE" in str(v) for k, v in data.items()
          if k != "_readme"))
check("no network-capable imports in gconfig/outputs/generate",
      not re.search(
          r"^\s*(import|from)\s+(googleapiclient|google|requests|urllib|"
          r"http\.client|socket)\b",
          "\n".join((REPO / "tools" / m).read_text(encoding="utf-8")
                    for m in ("gconfig.py", "outputs.py", "generate.py")),
          re.M))

print()
if failures:
    raise SystemExit(f"FAILED: {failures}")
print(f"All step-3.1 checks passed ({len(ics_uids)} events previewed).")
