#!/usr/bin/env python3
"""Run every regression suite and summarise.

Each verify_*.py in this directory is a self-contained suite named after
the build step that introduced it (PLAN.md). They run the real CLI against
samples/pat.sample.xlsx (dummy data) only, write outputs to temp dirs
outside the repo, and never touch a real workbook.

Usage:  python tools/tests/run_all.py       (exit 0 = all suites passed)
"""

import os
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).parent
suites = sorted(HERE.glob("verify_*.py"))
# The suites print em-dashes etc.; force UTF-8 stdout so they pass from
# any shell (Git Bash's default codepage crashes them otherwise).
ENV = dict(os.environ, PYTHONIOENCODING="utf-8")
failed = []
for suite in suites:
    print(f"=== {suite.name} ===", flush=True)
    result = subprocess.run([sys.executable, str(suite)], env=ENV)
    if result.returncode != 0:
        failed.append(suite.name)
    print(flush=True)

if failed:
    sys.exit(f"FAILED suites: {', '.join(failed)}")
print(f"All {len(suites)} suites passed.")
