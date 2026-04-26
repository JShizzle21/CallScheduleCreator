"""Test bootstrap: ensure both project root and src/ are on sys.path.

Internal modules live in `src/` (config, validation, data_bundle, etc.) to
keep the project root uncluttered for end users. The CLI entry point
(`scheduler_main.py`) stays at root. Tests need both directories on the
import path so existing flat-style imports (`from config import X`,
`import scheduler_main as sm`) keep working without per-test changes.

pytest auto-discovers this file before collecting tests in this directory.
"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "src")

for path in (ROOT, SRC):
    if path not in sys.path:
        sys.path.insert(0, path)
