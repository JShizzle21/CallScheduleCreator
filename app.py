"""Streamlit GUI for the Call Schedule Creator.

Phase 5a — scaffold + upload section only. Settings, Run, and Results are
stubbed. See docs/gui_plan.md for the full spec.

Runtime notes:
- Requires streamlit >= 1.38 (pinned in requirements.txt).
- Uploaded files land in a per-session temp directory created under the
  OS temp location via tempfile.mkdtemp(prefix="CallScheduler_"). On
  Windows this is typically C:\\Users\\<user>\\AppData\\Local\\Temp\\.
  Files in the session tmpdir are deleted when the user clicks
  "Reset session"; orphaned tmpdirs from prior sessions (older than
  24 hours) are cleaned on launch.
- All GUI state lives in st.session_state. Nothing persists to disk
  except via "Save as defaults" (Phase 5b), which overwrites config.yaml.
"""

from __future__ import annotations

import shutil
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

import streamlit as st
from openpyxl import load_workbook

from config import load_default_config

TMPDIR_PREFIX = "CallScheduler_"
ORPHAN_AGE_SECONDS = 24 * 3600


# ---------------------------------------------------------------------------
# Upload slot definitions
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class UploadSlot:
    key: str
    label: str
    saved_filename: str  # What we save the upload as inside the tmpdir.
    required: bool
    validator: Callable[[Path], Optional[str]]
    help: str


def _validate_flow(path: Path) -> Optional[str]:
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        return f"Could not open workbook: {exc}"
    if "master_block_calendar" not in wb.sheetnames:
        return (
            "Flow sheet error: expected a sheet named 'master_block_calendar'. "
            f"Found: {', '.join(wb.sheetnames)}"
        )
    ws = wb["master_block_calendar"]
    # Row 2, column B+ should hold block start dates.
    date_cells = [ws.cell(row=2, column=c).value for c in range(2, min(ws.max_column, 30) + 1)]
    if not any(hasattr(v, "year") for v in date_cells):
        return (
            "Flow sheet error: row 2 does not contain any parseable dates in "
            "columns B onward. Expected block start dates there."
        )
    return None


def _validate_headers(path: Path, required_headers: set[str], label: str) -> Optional[str]:
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        return f"Could not open workbook: {exc}"
    ws = wb.active
    headers = {
        str(ws.cell(row=1, column=c).value).strip().lower()
        for c in range(1, ws.max_column + 1)
        if ws.cell(row=1, column=c).value is not None
    }
    missing = {h.lower() for h in required_headers} - headers
    if missing:
        return (
            f"{label} error: header row is missing required column(s): "
            f"{', '.join(sorted(missing))}. Found: {', '.join(sorted(headers))}."
        )
    return None


def _validate_rotation_rules(path: Path) -> Optional[str]:
    return _validate_headers(path, {"rotation", "PGY1", "PGY2", "PGY3"}, "Rotation rules")


def _validate_no_call(path: Path) -> Optional[str]:
    return _validate_headers(path, {"name", "date"}, "No-call days")


def _validate_holidays(path: Path) -> Optional[str]:
    return _validate_headers(path, {"date"}, "Holidays")


def _validate_clinic(path: Path) -> Optional[str]:
    return _validate_headers(path, {"name", "date"}, "Clinic days")


def _validate_completed(path: Path) -> Optional[str]:
    return _validate_headers(path, {"date"}, "Completed calls")


UPLOAD_SLOTS: list[UploadSlot] = [
    UploadSlot(
        key="flow",
        label="Rotation flow sheet",
        saved_filename="flow.xlsx",
        required=True,
        validator=_validate_flow,
        help="Must contain a sheet named 'master_block_calendar' with block start dates in row 2.",
    ),
    UploadSlot(
        key="rotation_rules",
        label="Rotation rules",
        saved_filename="rotation_rules.xlsx",
        required=True,
        validator=_validate_rotation_rules,
        help="Columns: rotation, PGY1, PGY2, PGY3. Values per row: ELIGIBLE, AVOID, or NO_CALL.",
    ),
    UploadSlot(
        key="no_call_days",
        label="No-call days",
        saved_filename="no_call_days.xlsx",
        required=True,
        validator=_validate_no_call,
        help="Columns: name, date. May be empty if no residents have no-call requests.",
    ),
    UploadSlot(
        key="holidays",
        label="Holidays",
        saved_filename="holidays.xlsx",
        required=True,
        validator=_validate_holidays,
        help="Columns: date, name, upper, intern. May be empty if there are no holiday pre-assignments.",
    ),
    UploadSlot(
        key="clinic_days",
        label="Clinic days",
        saved_filename="clinic_days.xlsx",
        required=False,
        validator=_validate_clinic,
        help="Optional. Columns: name, date.",
    ),
    UploadSlot(
        key="completed_calls",
        label="Completed calls",
        saved_filename="completed_calls.xlsx",
        required=False,
        validator=_validate_completed,
        help="Optional. Columns: date, upper, intern. Required only in partial-year mode.",
    ),
]


# ---------------------------------------------------------------------------
# Session-state management
# ---------------------------------------------------------------------------


def _cleanup_orphan_tmpdirs() -> None:
    """Delete CallScheduler_* tmpdirs older than 24 hours.

    Streamlit has no reliable shutdown hook, so prior sessions can leave
    tmpdirs behind. We only touch entries older than ORPHAN_AGE_SECONDS
    so concurrent sessions can't delete each other's state.
    """
    root = Path(tempfile.gettempdir())
    now = time.time()
    for child in root.glob(f"{TMPDIR_PREFIX}*"):
        try:
            if child.is_dir() and now - child.stat().st_mtime > ORPHAN_AGE_SECONDS:
                shutil.rmtree(child, ignore_errors=True)
        except OSError:
            pass


def _init_session_state() -> None:
    if "initialized" in st.session_state:
        return
    _cleanup_orphan_tmpdirs()
    st.session_state.tmpdir = Path(tempfile.mkdtemp(prefix=TMPDIR_PREFIX))
    # uploads[slot.key] = {"saved_path": Path, "original_name": str,
    #                       "uploaded_at": float, "error": Optional[str]}
    st.session_state.uploads = {}
    # Seed config once. Phase 5b will wire widgets to these values.
    config, _paths = load_default_config()
    st.session_state.config = config
    st.session_state.initialized = True


def _reset_session() -> None:
    """Delete the tmpdir and clear session state so the page starts fresh."""
    tmpdir = st.session_state.get("tmpdir")
    if tmpdir is not None:
        shutil.rmtree(tmpdir, ignore_errors=True)
    for key in list(st.session_state.keys()):
        del st.session_state[key]


# ---------------------------------------------------------------------------
# Upload section
# ---------------------------------------------------------------------------


def _save_upload(slot: UploadSlot, uploaded_file) -> None:
    """Persist the Streamlit UploadedFile to the session tmpdir and validate."""
    dest = st.session_state.tmpdir / slot.saved_filename
    with open(dest, "wb") as fh:
        fh.write(uploaded_file.getbuffer())
    error = slot.validator(dest)
    st.session_state.uploads[slot.key] = {
        "saved_path": dest,
        "original_name": uploaded_file.name,
        "uploaded_at": time.time(),
        "error": error,
    }


def _remove_upload(slot_key: str) -> None:
    entry = st.session_state.uploads.pop(slot_key, None)
    if entry is not None:
        try:
            Path(entry["saved_path"]).unlink(missing_ok=True)
        except OSError:
            pass


def _render_upload_slot(slot: UploadSlot) -> None:
    required_tag = "" if slot.required else " _(optional)_"
    st.markdown(f"**{slot.label}**{required_tag}")
    st.caption(slot.help)

    entry = st.session_state.uploads.get(slot.key)
    if entry is not None:
        cols = st.columns([4, 1])
        if entry["error"] is None:
            cols[0].success(
                f"✓ {entry['original_name']} — saved as `{slot.saved_filename}`"
            )
        else:
            cols[0].error(f"✗ {entry['original_name']} — {entry['error']}")
        if cols[1].button("Remove", key=f"remove_{slot.key}"):
            _remove_upload(slot.key)
            st.rerun()
    else:
        st.caption("No file loaded.")

    uploaded = st.file_uploader(
        f"Upload {slot.label}",
        type=["xlsx"],
        key=f"uploader_{slot.key}",
        label_visibility="collapsed",
    )
    if uploaded is not None:
        existing = entry is not None and entry["original_name"] == uploaded.name
        if not existing:
            _save_upload(slot, uploaded)
            st.rerun()

    st.divider()


def _all_required_valid() -> bool:
    for slot in UPLOAD_SLOTS:
        if not slot.required:
            continue
        entry = st.session_state.uploads.get(slot.key)
        if entry is None or entry["error"] is not None:
            return False
    return True


# ---------------------------------------------------------------------------
# Page
# ---------------------------------------------------------------------------


def main() -> None:
    st.set_page_config(page_title="Call Schedule Creator", layout="wide")
    _init_session_state()

    header_cols = st.columns([6, 1])
    header_cols[0].title("Call Schedule Creator")
    if header_cols[1].button("Reset session"):
        _reset_session()
        st.rerun()

    with st.expander("1. Upload input files", expanded=True):
        for slot in UPLOAD_SLOTS:
            _render_upload_slot(slot)

    with st.expander("2. Settings", expanded=False):
        st.info("Settings panel — coming in Phase 5b.")

    with st.expander("3. Run schedule", expanded=False):
        ready = _all_required_valid()
        if not ready:
            st.warning("Upload all required files above before running.")
        st.button("Run (coming in Phase 5c)", disabled=True)

    with st.expander("4. Results", expanded=False):
        st.caption("Results will appear here after a run completes — Phase 5c.")


if __name__ == "__main__":
    main()
