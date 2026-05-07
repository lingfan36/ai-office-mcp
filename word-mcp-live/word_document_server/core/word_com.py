"""COM connection manager and snapshot-based undo for Microsoft Word on Windows."""

import json
import os
import shutil
import sys
import time
import unicodedata
import uuid
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

SNAPSHOT_DIR = Path(os.environ.get("WORD_SNAPSHOT_DIR", Path.home() / ".word_mcp" / "snapshots"))
MAX_SNAPSHOTS = int(os.environ.get("WORD_MAX_SNAPSHOTS", "20"))

_app = None  # singleton Word.Application COM object


def reset_word_app():
    """Clear the cached COM app reference (forces reconnect on next call)."""
    global _app
    _app = None


def _docs_accessible(app) -> bool:
    try:
        _ = app.Documents.Count
        return True
    except Exception:
        return False


def get_word_app(visible: bool = True, auto_launch: bool = True):
    """Return the running Word.Application, launching one if needed.

    Args:
        visible: Make Word window visible (default True).
        auto_launch: If Word is not running, launch it (default True).

    Returns:
        Word.Application COM object.

    Raises:
        RuntimeError: If not on Windows, or Word cannot be found/launched.
    """
    if sys.platform != "win32":
        raise RuntimeError("Word COM automation is only available on Windows")

    import win32com.client
    import pywintypes

    global _app

    # Check cached singleton
    try:
        if _app is not None:
            _ = _app.Version  # liveness probe
            if _docs_accessible(_app):
                return _app
            _app = None
    except Exception:
        _app = None

    # Try a running instance via GetActiveObject
    try:
        candidate = win32com.client.GetActiveObject("Word.Application")
        if _docs_accessible(candidate):
            _app = candidate
            return _app
        # Scan ROT for an instance with open documents
        found = _find_word_with_docs()
        if found is not None:
            _app = found
            return _app
        # Use the empty instance (caller can open a file)
        _app = candidate
        return _app
    except pywintypes.com_error:
        pass
    except Exception:
        pass

    # Try ROT scan
    found = _find_word_with_docs()
    if found is not None:
        _app = found
        return _app

    # Auto-launch Word if allowed
    if not auto_launch:
        raise RuntimeError(
            "Microsoft Word is not running. Open Word first or set auto_launch=True."
        )

    _app = win32com.client.Dispatch("Word.Application")
    try:
        _app.Visible = visible
    except Exception:
        pass
    # Wait up to 6 s for Word to finish initialising
    for _ in range(12):
        if _docs_accessible(_app):
            return _app
        time.sleep(0.5)
    return _app


def _find_word_with_docs():
    """Scan the Running Object Table for a Word.Application with open documents."""
    try:
        import pythoncom
        import win32com.client

        rot = pythoncom.GetRunningObjectTable(0)
        enum = rot.EnumRunning()
        monikers_to_retry = []

        while True:
            batch = enum.Next(1)
            if not batch:
                break
            moniker = batch[0]
            try:
                ctx = pythoncom.CreateBindCtx(0)
                name = moniker.GetDisplayName(ctx, None)
                obj = rot.GetObject(moniker)
                dispatch = obj.QueryInterface(pythoncom.IID_IDispatch)
                com_obj = win32com.client.Dispatch(dispatch)
                if hasattr(com_obj, "Documents") and hasattr(com_obj, "ActiveDocument"):
                    if com_obj.Documents.Count > 0:
                        return com_obj
                if name and (name.lower().endswith(".docx") or name.lower().endswith(".doc")):
                    monikers_to_retry.append((name, moniker))
            except Exception:
                try:
                    ctx = pythoncom.CreateBindCtx(0)
                    name = moniker.GetDisplayName(ctx, None)
                    if name and (name.lower().endswith(".docx") or name.lower().endswith(".doc")):
                        monikers_to_retry.append((name, moniker))
                except Exception:
                    pass
                continue

        for name, moniker in monikers_to_retry:
            try:
                obj = rot.GetObject(moniker)
                dispatch = obj.QueryInterface(pythoncom.IID_IDispatch)
                doc = win32com.client.Dispatch(dispatch)
                app = doc.Application
                if app.Documents.Count > 0:
                    return app
            except Exception:
                continue
    except Exception:
        pass
    return None


def find_document(app, filename: str = None):
    """Find an open document by filename.

    Args:
        app: Word.Application COM object.
        filename: Document name (basename) or full path.
                  If None or empty, returns the active document.

    Returns:
        Document COM object.
    """
    if app.Documents.Count == 0:
        raise ValueError("No documents are open in Word")

    if not filename:
        return app.ActiveDocument

    target_basename = unicodedata.normalize("NFC", os.path.basename(filename)).lower()
    target_fullpath = (
        unicodedata.normalize("NFC", os.path.normpath(filename)).lower()
        if os.path.isabs(filename)
        else None
    )

    for i in range(1, app.Documents.Count + 1):
        doc = app.Documents(i)
        if unicodedata.normalize("NFC", doc.Name).lower() == target_basename:
            return doc
        if (
            target_fullpath
            and unicodedata.normalize("NFC", os.path.normpath(doc.FullName)).lower()
            == target_fullpath
        ):
            return doc

    open_docs = [app.Documents(i).Name for i in range(1, app.Documents.Count + 1)]
    raise ValueError(
        f"Document '{filename}' is not open in Word. Open documents: {open_docs}"
    )


@contextmanager
def undo_record(app, name: str):
    """Wrap COM mutations in a single Word UndoRecord (one Ctrl+Z entry)."""
    rec = None
    try:
        rec = app.UndoRecord
        if rec.IsRecordingCustomRecord:
            try:
                rec.EndCustomRecord()
            except Exception:
                pass
        rec.StartCustomRecord(name[:64])
    except Exception:
        rec = None
    try:
        yield
    finally:
        if rec is not None:
            try:
                rec.EndCustomRecord()
            except Exception:
                pass


# ---------------------------------------------------------------------------
# Snapshot-based undo
# ---------------------------------------------------------------------------

class SnapshotUndo:
    """Per-document snapshot stack for AI-safe undo, persisted to disk."""

    _stack: list[dict] = []
    _loaded: bool = False
    _INDEX_FILE: Path = SNAPSHOT_DIR / "undo_index.json"

    @classmethod
    def _ensure_loaded(cls):
        if cls._loaded:
            return
        cls._loaded = True
        if cls._INDEX_FILE.exists():
            try:
                data = json.loads(cls._INDEX_FILE.read_text(encoding="utf-8"))
                cls._stack = [e for e in data if Path(e["snapshot"]).exists()]
            except Exception:
                cls._stack = []

    @classmethod
    def _save_index(cls):
        try:
            SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
            cls._INDEX_FILE.write_text(
                json.dumps(cls._stack, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            pass

    @classmethod
    def push_path(cls, doc_path: str, doc_name: str) -> str:
        """Save a snapshot by file path. Works on any OS (no COM required)."""
        cls._ensure_loaded()
        SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
        snap_id = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}"
        ext = os.path.splitext(doc_path)[1] or ".docx"
        snap_path = str(SNAPSHOT_DIR / f"{snap_id}{ext}")
        try:
            shutil.copy2(doc_path, snap_path)
        except Exception as e:
            return f"snapshot_failed:{e}"
        cls._stack.append({
            "id": snap_id,
            "doc_name": doc_name,
            "doc_path": doc_path,
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
    def pop_path(cls, doc_name: str = None) -> dict:
        """Restore most recent snapshot by file path only (no COM app needed — macOS)."""
        cls._ensure_loaded()
        if not cls._stack:
            return {"success": False, "message": "Undo stack is empty"}

        idx = None
        for i in range(len(cls._stack) - 1, -1, -1):
            entry = cls._stack[i]
            if doc_name is None or entry["doc_name"].lower() == doc_name.lower():
                idx = i
                break

        if idx is None:
            return {"success": False, "message": f"No snapshot for document: {doc_name!r}"}

        entry = cls._stack.pop(idx)
        snap_path = entry["snapshot"]
        doc_path = entry["doc_path"]

        try:
            shutil.copy2(snap_path, doc_path)
        except Exception as e:
            cls._stack.insert(idx, entry)
            return {"success": False, "message": f"Failed to restore: {e}"}
        try:
            os.remove(snap_path)
        except OSError:
            pass
        cls._save_index()
        return {"success": True, "restored": doc_path, "snapshot_id": entry["id"]}

    @classmethod
    def push(cls, doc) -> str:
        """Save a snapshot of *doc* before a destructive operation. Returns snapshot id."""
        cls._ensure_loaded()
        SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
        snap_id = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}"

        try:
            src = doc.FullName
            if not src or not os.path.exists(src):
                # Unsaved document — save to temp first
                tmp = str(SNAPSHOT_DIR / f"_tmp_{snap_id}.docx")
                doc.SaveAs2(tmp, FileFormat=16)  # wdFormatXMLDocument=16
                snap_path = str(SNAPSHOT_DIR / f"{snap_id}.docx")
                shutil.copy2(tmp, snap_path)
                os.remove(tmp)
            else:
                ext = os.path.splitext(src)[1] or ".docx"
                snap_path = str(SNAPSHOT_DIR / f"{snap_id}{ext}")
                shutil.copy2(src, snap_path)
        except Exception as e:
            return f"snapshot_failed:{e}"

        cls._stack.append({
            "id": snap_id,
            "doc_name": doc.Name,
            "doc_path": doc.FullName,
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
    def pop(cls, app, doc_name: str = None) -> dict:
        """Restore the most recent snapshot for *doc_name* (Windows COM path)."""
        cls._ensure_loaded()
        if not cls._stack:
            return {"success": False, "message": "Undo stack is empty"}

        idx = None
        for i in range(len(cls._stack) - 1, -1, -1):
            entry = cls._stack[i]
            if doc_name is None or entry["doc_name"].lower() == doc_name.lower():
                idx = i
                break

        if idx is None:
            return {"success": False, "message": f"No snapshot for document: {doc_name!r}"}

        entry = cls._stack.pop(idx)
        doc_path = entry["doc_path"]
        snap_path = entry["snapshot"]

        try:
            for i in range(1, app.Documents.Count + 1):
                d = app.Documents(i)
                if os.path.normpath(d.FullName).lower() == os.path.normpath(doc_path).lower():
                    d.Close(SaveChanges=False)
                    break
        except Exception:
            pass

        shutil.copy2(snap_path, doc_path)
        try:
            os.remove(snap_path)
        except OSError:
            pass
        cls._save_index()
        app.Documents.Open(doc_path)
        return {"success": True, "restored": doc_path, "snapshot_id": entry["id"]}

    @classmethod
    def list(cls, doc_name: str = None) -> list[dict]:
        cls._ensure_loaded()
        snapshots = cls._stack
        if doc_name:
            snapshots = [s for s in snapshots if s["doc_name"].lower() == doc_name.lower()]
        return [
            {"id": s["id"], "doc_name": s["doc_name"], "timestamp": s["id"][:15]}
            for s in snapshots
        ]

    @classmethod
    def clear(cls, doc_name: str = None):
        cls._ensure_loaded()
        if doc_name:
            keep, remove = [], []
            for s in cls._stack:
                (remove if s["doc_name"].lower() == doc_name.lower() else keep).append(s)
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
