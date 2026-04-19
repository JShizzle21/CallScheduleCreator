"""Tests for config.save_config — the "Save as defaults" path for the GUI.

The contract: save_config updates only the keys passed in, preserves every
other key, and preserves comments on keys that were not touched. The GUI
will use this to overwrite config.yaml from session_state without
destroying the file's annotations.
"""

from __future__ import annotations

import pytest

from config import load_config, save_config


def _write(path, text: str) -> None:
    path.write_text(text, encoding="utf-8")


def test_save_config_updates_values_and_preserves_unrelated_keys(tmp_path):
    cfg = tmp_path / "config.yaml"
    _write(
        cfg,
        "# header comment\n"
        "FAIRNESS_GAP_WEIGHT: 3.0\n"
        "SPACING_WEIGHT: 1.0\n"
        "FLOW_XLSX: data/flow.xlsx\n",
    )

    save_config({"FAIRNESS_GAP_WEIGHT": 5.5}, path=str(cfg))

    loaded = load_config(str(cfg))
    assert loaded["FAIRNESS_GAP_WEIGHT"] == 5.5
    assert loaded["SPACING_WEIGHT"] == 1.0
    assert loaded["FLOW_XLSX"] == "data/flow.xlsx"


def test_save_config_preserves_comments(tmp_path):
    cfg = tmp_path / "config.yaml"
    original = (
        "# This comment must survive round-trip.\n"
        "\n"
        "# Inline section header\n"
        "FAIRNESS_GAP_WEIGHT: 3.0  # weight for pool-min gap\n"
        "SPACING_WEIGHT: 1.0\n"
    )
    _write(cfg, original)

    save_config({"SPACING_WEIGHT": 2.25}, path=str(cfg))

    text = cfg.read_text(encoding="utf-8")
    assert "# This comment must survive round-trip." in text
    assert "# Inline section header" in text
    assert "# weight for pool-min gap" in text
    assert "SPACING_WEIGHT: 2.25" in text


def test_save_config_missing_file_raises(tmp_path):
    missing = tmp_path / "does_not_exist.yaml"
    with pytest.raises(FileNotFoundError):
        save_config({"FAIRNESS_GAP_WEIGHT": 1.0}, path=str(missing))


def test_save_config_ignores_extra_keys_not_in_file(tmp_path):
    """save_config adds new keys — callers are trusted to pass valid keys.

    This documents the current behavior: unknown keys get appended to the
    file rather than silently dropped. The GUI filters its session_state
    keys before calling save_config.
    """
    cfg = tmp_path / "config.yaml"
    _write(cfg, "EXISTING: 1\n")

    save_config({"EXISTING": 2, "NEW_KEY": "hello"}, path=str(cfg))

    loaded = load_config(str(cfg))
    assert loaded["EXISTING"] == 2
    assert loaded["NEW_KEY"] == "hello"
