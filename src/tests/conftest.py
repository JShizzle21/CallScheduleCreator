"""Test bootstrap: ensure both project root and src/ are on sys.path.

After the May refactor this file sits at src/tests/conftest.py — so the
project root is two directories up (../../ from here). Internal modules
live in src/ (config, validation, data_bundle, scheduler_main, etc.);
adding src/ to sys.path keeps existing flat-style imports working
(`from config import X`, `import scheduler_main as sm`) without
touching each test file.

pytest auto-discovers this file before collecting tests in this directory.
"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SRC = os.path.join(ROOT, "src")

for path in (ROOT, SRC):
    if path not in sys.path:
        sys.path.insert(0, path)
