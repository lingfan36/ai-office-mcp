"""COM connection management and snapshot-based undo for Excel."""
import json
import os
import shutil
import time
import uuid
from datetime import datetime
from pathlib import Path

import win32com.client
import pywintypes

SNAPSHOT_DIR = Path(os.environ.get("EXCEL_SNAPSHOT_DIR", Path.home() / ".excel_mcp" / "snapshots"))
MAX_SNAPSHOTS = int(os.environ.get("EXCEL_MAX_SNAPSHOTS", "20"))
_INDEX_FILE = SNAPSHOT_DIR / "undo_index.json"

_app = None  # singleton Excel.Application COM object


def reset_excel_app():
    """Clear the cached COM app reference (forces reconnect on next call)."""
    global _app
    _app = None


def _workbooks_accessible(app) -> bool:
    """Return True if app.Workbooks is accessible (not blocked by WPS home screen)."""
    try:
        _ = app.Workbooks.Count
        return True
    except Exception:
        return False


def _dismiss_blocking_dialogs():
    """Close WPS NUIDialog modals and stray 'Open With' windows that block COM."""
    try:
        import win32gui, win32con
        def _cb(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return True
            cls = win32gui.GetClassName(hwnd)
            if cls in ("NUIDialog", "Open With Dummy Window Class For Interim Dialog"):
                try:
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                except Exception:
                    pass
            return True
        win32gui.EnumWindows(_cb, None)
    except Exception:
        pass


def _suppress_alerts(app):
    """Set DisplayAlerts=False to prevent WPS from blocking on format/save dialogs."""
    try:
        app.DisplayAlerts = False
    except Exception:
        pass


def get_excel_app(visible: bool = True):
    """Return the running Excel.Application, launching one if needed."""
    global _app
    # Proactively dismiss any WPS blocking dialogs before every call
    _dismiss_blocking_dialogs()
    try:
        if _app is not None:
            _ = _app.Version  # liveness probe
            if _workbooks_accessible(_app):
                return _app
            # Workbooks blocked — dismiss and retry
            _dismiss_blocking_dialogs()
            time.sleep(0.3)
            if _workbooks_accessible(_app):
                return _app
            _app = None
    except Exception:
        _app = None

    # Try a running instance
    try:
        candidate = win32com.client.GetActiveObject("Excel.Application")
        if _workbooks_accessible(candidate):
            _app = candidate
            return _app
        # Blocked — dismiss dialogs and retry
        _dismiss_blocking_dialogs()
        time.sleep(0.5)
        if _workbooks_accessible(candidate):
            _app = candidate
            return _app
        # Still blocked — fall through to fresh Dispatch
    except pywintypes.com_error:
        pass

    _app = win32com.client.Dispatch("Excel.Application")
    try:
        _app.Visible = visible
    except Exception:
        pass  # WPS may not support Visible in automation mode
    # Wait up to 6 s, dismissing dialogs along the way
    for _ in range(12):
        if _workbooks_accessible(_app):
            return _app
        _dismiss_blocking_dialogs()
        time.sleep(0.5)
    return _app  # return even if still not fully ready


def find_workbook(app, name_or_path: str = None):
    """Return a Workbook by name/path, or the active workbook when None."""
    if not name_or_path:
        wb = app.ActiveWorkbook
        if wb is None:
            raise ValueError("No active workbook. Open a file first.")
        return wb
    target = name_or_path.lower()
    for wb in app.Workbooks:
        if wb.Name.lower() == target or wb.FullName.lower() == target:
            return wb
    raise ValueError(f"Workbook not found: {name_or_path!r}")


def find_sheet(wb, sheet_name: str = None):
    """Return a Worksheet by name, or the active sheet when None."""
    if not sheet_name:
        return wb.ActiveSheet
    for ws in wb.Worksheets:
        if ws.Name == sheet_name:
            return ws
    raise ValueError(f"Sheet not found: {sheet_name!r}")


# ---------------------------------------------------------------------------
# Snapshot-based undo
# ---------------------------------------------------------------------------

class SnapshotUndo:
    """Per-workbook snapshot stack with disk-persisted index (survives server restarts)."""

    _stack: list[dict] = []
    _loaded: bool = False

    # ── index persistence ──────────────────────────────────────────────────────

    @classmethod
    def _ensure_loaded(cls):
        if cls._loaded:
            return
        cls._loaded = True
        if _INDEX_FILE.exists():
            try:
                data = json.loads(_INDEX_FILE.read_text(encoding="utf-8"))
                # Discard entries whose snapshot files no longer exist (orphans)
                cls._stack = [e for e in data if Path(e["snapshot"]).exists()]
            except Exception:
                cls._stack = []

    @classmethod
    def _save_index(cls):
        try:
            SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
            _INDEX_FILE.write_text(
                json.dumps(cls._stack, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
        except Exception:
            pass

    # ── public API ─────────────────────────────────────────────────────────────

    @classmethod
    def push(cls, wb) -> str:
        """Save a snapshot of *wb* before a destructive operation. Returns snapshot id."""
        cls._ensure_loaded()
        SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
        snap_id = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}"
        snap_path = str(SNAPSHOT_DIR / f"{snap_id}.xlsx")
        wb.SaveCopyAs(snap_path)
        cls._stack.append({
            "id": snap_id,
            "wb_name": wb.Name,
            "wb_path": wb.FullName,
            "snapshot": snap_path,
        })
        while len(cls._stack) > MAX_SNAPSHOTS:
            old = cls._stack.pop(0)
            try:
                os.remove(old["snapshot"])
            except OSError:
                pass
        cls._save_index()
        return snap_id

    @classmethod
    def pop(cls, app, workbook_name: str = None) -> dict:
        """Restore the most recent snapshot for *workbook_name*."""
        cls._ensure_loaded()
        if not cls._stack:
            return {"success": False, "message": "Undo stack is empty"}

        idx = None
        for i in range(len(cls._stack) - 1, -1, -1):
            entry = cls._stack[i]
            if workbook_name is None or entry["wb_name"].lower() == workbook_name.lower():
                idx = i
                break

        if idx is None:
            return {"success": False, "message": f"No snapshot for workbook: {workbook_name!r}"}

        entry = cls._stack.pop(idx)
        wb_path = entry["wb_path"]
        snap_path = entry["snapshot"]

        try:
            for wb in app.Workbooks:
                if wb.FullName == wb_path:
                    wb.Close(SaveChanges=False)
                    break
        except Exception:
            pass

        shutil.copy2(snap_path, wb_path)
        try:
            os.remove(snap_path)
        except OSError:
            pass
        cls._save_index()
        app.Workbooks.Open(wb_path)
        return {"success": True, "restored": wb_path, "snapshot_id": entry["id"]}

    @classmethod
    def list(cls, workbook_name: str = None) -> list[dict]:
        cls._ensure_loaded()
        snapshots = cls._stack
        if workbook_name:
            snapshots = [s for s in snapshots if s["wb_name"].lower() == workbook_name.lower()]
        return [{"id": s["id"], "wb_name": s["wb_name"], "timestamp": s["id"][:15]} for s in snapshots]

    @classmethod
    def clear(cls, workbook_name: str = None):
        cls._ensure_loaded()
        if workbook_name:
            keep, remove = [], []
            for s in cls._stack:
                (remove if s["wb_name"].lower() == workbook_name.lower() else keep).append(s)
            cls._stack = keep
        else:
            remove = list(cls._stack)
            cls._stack = []
        for s in remove:
            try:
                os.remove(s["snapshot"])
            except OSError:
                pass
        cls._save_index()
