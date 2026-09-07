"""Regression suite from build step 2.3: AI augmentation (AC-V4, AC-V5).

Sample consenting rows (ai_visible? = Y, set in make_template.py):
  banks #2               Example Bank plc — Savings account (Pat's)
  building_societies #1  Sampletown Building Society — Instant access saver
Everything else is blank/N and must never reach a query.
"""

import hashlib
import json
import subprocess
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SAMPLE = REPO / "samples" / "pat.sample.xlsx"
SCRATCH = Path(tempfile.mkdtemp(prefix="crisisplan-verify23-"))
AUGMENT = REPO / "tools" / "augment.py"
CLI = REPO / "tools" / "generate.py"
sys.path.insert(0, str(REPO / "tools"))

failures = []


def check(name, ok, detail=""):
    print(f"  [{'PASS' if ok else 'FAIL'}] {name}" + (f" — {detail}" if detail else ""))
    if not ok:
        failures.append(name)


def run(script, *args, stdin=None):
    return subprocess.run(
        [sys.executable, str(script), *[str(a) for a in args]],
        input=stdin, capture_output=True, text=True, encoding="utf-8",
        errors="replace")


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def run_cli(out_dir, cache):
    r = run(CLI, "-f", SAMPLE, "--date", "2030-01-15", "--out", out_dir,
            "--augment-cache", cache)
    if r.returncode != 0:
        print(r.stdout, r.stderr)
        raise SystemExit("CLI failed")
    return r.stdout


BANKS_KEY = "banks/example-bank-plc-savings-account"
BS_KEY = "building_societies/sampletown-building-society-instant-access-saver"

print("AC-V4: consent gate prints minimised queries only (--list)")
r = run(AUGMENT, "-f", SAMPLE, "--list")
check("--list exits 0", r.returncode == 0, r.stderr.strip())
out = r.stdout
check("both consenting institutions listed",
      "Example Bank plc" in out and "Sampletown Building Society" in out)
check("rows identified only as tab #id",
      "banks #2" in out and "building_societies #1" in out)
leaks = ["Pat", "Alex", "Savings account", "Instant access", "Passbook",
         "Password manager", "home safe"]
check("no owner / account label / location in the gate output",
      not any(term in out for term in leaks),
      ", ".join(t for t in leaks if t in out))
non_consenting = ["Another Bank", "Example Investments",
                  "Example Registrars", "Example Pensions",
                  "Example Insurance", "example.com"]
check("no non-consenting row's institution appears",
      not any(term in out for term in non_consenting),
      ", ".join(t for t in non_consenting if t in out))

print("AC-V4: explicit confirmation required before any lookup")
cache = SCRATCH / "cache.json"
r = run(AUGMENT, "-f", SAMPLE, "--cache", cache, stdin="no\n")
check("answering 'no' aborts non-zero", r.returncode != 0)
check("aborted run writes no cache", not cache.exists())

print("AC-V5: stub run — cache round-trip, workbook untouched")
sample_hash = sha256(SAMPLE)
r = run(AUGMENT, "-f", SAMPLE, "--cache", cache, stdin="yes\n")
check("confirmed stub run exits 0", r.returncode == 0, r.stderr.strip())
check("cache file written", cache.exists())
entries = json.loads(cache.read_text(encoding="utf-8"))["entries"]
check("cache keyed by the two consenting planKeys",
      set(entries) == {BANKS_KEY, BS_KEY}, ", ".join(entries))
check("every cached row has lookup lines",
      all(e["lines"] for e in entries.values()))
check("workbook hash unchanged (never written)",
      sha256(SAMPLE) == sample_hash)

print("AC-V5: renderers merge the cache; .ics never carries it")
with_dir, without_dir = SCRATCH / "run-with", SCRATCH / "run-without"
stdout = run_cli(with_dir, cache)
check("health summary reports the merge",
      "augment cache" in stdout and "2 row(s)" in stdout)
md = (with_dir / "checklist.md").read_text(encoding="utf-8")
check("checklist carries both rows' [ai] lines",
      "[ai] [ai-stub] Placeholder for 'Example Bank plc'" in md
      and "[ai] [ai-stub] Placeholder for 'Sampletown Building Society'"
      in md)
check("each row's lines merged exactly once in the master",
      md.count("[ai-stub] Placeholder") == 2)
ics = (with_dir / "plan.ics").read_bytes().decode("utf-8")
check(".ics has no augment lines (external surface)",
      "ai-stub" not in ics and "[ai]" not in ics)
alex = (with_dir / "checklist-alex-sample.md").read_text(encoding="utf-8")
check("assignee's personal checklist carries the lines too",
      "[ai-stub] Placeholder for 'Example Bank plc'" in alex)

print("AC-V5: outputs unchanged when the cache is absent")
stdout = run_cli(without_dir, SCRATCH / "no-such-cache.json")
check("health summary reports no cache", "no cache" in stdout)
md_without = (without_dir / "checklist.md").read_text(encoding="utf-8")
check("no [ai] lines without a cache", "[ai]" not in md_without)
stripped = [ln for ln in md.splitlines() if "[ai]" not in ln]
check("with-cache checklist minus [ai] lines == without-cache checklist",
      stripped == md_without.splitlines())
check(".ics identical with and without the cache",
      (with_dir / "plan.ics").read_bytes()
      == (without_dir / "plan.ics").read_bytes())

print("Manual paste flow (documented provider)")
cache2 = SCRATCH / "cache-manual.json"
pasted = "yes\nBereavement line: 0800 111 222\nhttps://example.org/deceased\n\n"
r = run(AUGMENT, "-f", SAMPLE, "--provider", "manual", "--cache", cache2,
        stdin=pasted)
check("manual run exits 0", r.returncode == 0, r.stderr.strip())
entries2 = json.loads(cache2.read_text(encoding="utf-8"))["entries"]
check("pasted lines cached under the first row's planKey",
      list(entries2) == [BANKS_KEY]
      and entries2[BANKS_KEY]["lines"]
      == ["Bereavement line: 0800 111 222", "https://example.org/deceased"])
check("nothing-pasted row skipped, visibly",
      "skipped (no lines): building_societies #1" in r.stdout)

print("Corrupt cache is ignored, run still succeeds")
bad = SCRATCH / "corrupt.json"
bad.write_text("not json at all", encoding="utf-8")
stdout = run_cli(SCRATCH / "run-corrupt", bad)
check("run OK with 'empty or unreadable' note",
      "empty or unreadable" in stdout)
check("no [ai] lines from a corrupt cache",
      "[ai]" not in (SCRATCH / "run-corrupt" / "checklist.md")
      .read_text(encoding="utf-8"))

print("Cache location & guards")
from augment import default_cache_path  # noqa: E402

check("in-repo workbook -> cache under gitignored private/",
      default_cache_path(SAMPLE)
      == REPO / "private" / "augment" / "pat.sample.augment.json")
check("outside-repo workbook -> sidecar beside it",
      default_cache_path(SCRATCH / "pat.xlsx")
      == SCRATCH / "pat.augment.json")
r = run(AUGMENT, "-f", SAMPLE, "--cache", REPO / "tools" / "cache.json")
check("tracked-tree cache path refused",
      r.returncode != 0 and "REFUSED" in r.stderr)

print()
if failures:
    raise SystemExit(f"FAILED: {failures}")
print(f"All step-2.3 checks passed ({len(entries)} rows cached).")
