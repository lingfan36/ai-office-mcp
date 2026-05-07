"""Snapshot-based undo tools for Word MCP.

Before any destructive operation, call word_take_snapshot to save the document
state. If the result is not satisfactory, call word_undo_last to restore.

Snapshots are file copies persisted to disk — they survive MCP server restarts
and work on Windows (via COM) and macOS (via JXA).
"""

import json
import os
import sys


async def word_take_snapshot(filename: str = None) -> str:
    """Save a snapshot of an open Word document for later undo.

    Call this BEFORE making changes you may want to reverse. Snapshots are
    persisted to disk and survive MCP server restarts.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with snapshot_id and confirmation.
    """
    if sys.platform == "win32":
        try:
            from word_document_server.core.word_com import get_word_app, find_document, SnapshotUndo

            app = get_word_app()
            doc = find_document(app, filename)
            snap_id = SnapshotUndo.push(doc)

            if snap_id.startswith("snapshot_failed:"):
                return json.dumps({
                    "success": False,
                    "error": snap_id[len("snapshot_failed:"):],
                    "hint": "Save the document to disk first, then retry.",
                })

            return json.dumps({
                "success": True,
                "snapshot_id": snap_id,
                "document": doc.Name,
                "total_snapshots": len(SnapshotUndo.list()),
            })
        except Exception as e:
            return json.dumps({"error": str(e)})

    elif sys.platform == "darwin":
        try:
            from word_document_server.core.word_mac import find_document as mac_find_document, get_word_app as mac_get_word_app
            from word_document_server.core.word_com import SnapshotUndo

            mac_get_word_app()
            doc_info = mac_find_document(None, filename)
            doc_path = doc_info.get("path", "")
            doc_name = doc_info.get("name", "")

            if not doc_path or not os.path.exists(doc_path):
                return json.dumps({
                    "success": False,
                    "error": "Document must be saved to disk before taking a snapshot.",
                })

            snap_id = SnapshotUndo.push_path(doc_path, doc_name)

            if snap_id.startswith("snapshot_failed:"):
                return json.dumps({
                    "success": False,
                    "error": snap_id[len("snapshot_failed:"):],
                })

            return json.dumps({
                "success": True,
                "snapshot_id": snap_id,
                "document": doc_name,
                "total_snapshots": len(SnapshotUndo.list()),
            })
        except Exception as e:
            return json.dumps({"error": str(e)})

    else:
        return json.dumps({"error": "Word automation is only supported on Windows and macOS"})


async def word_undo_last(filename: str = None) -> str:
    """Restore an open Word document to its most recent snapshot state.

    This closes the current document and reopens the snapshot copy, discarding
    all changes made since the snapshot was taken.

    Args:
        filename: Document name or path (None = active document). Used to select
                  which document's snapshot to restore.

    Returns:
        JSON confirming restoration or an error if no snapshot exists.
    """
    if sys.platform == "win32":
        try:
            from word_document_server.core.word_com import get_word_app, find_document, SnapshotUndo

            app = get_word_app()
            doc_name = None
            try:
                doc = find_document(app, filename)
                doc_name = doc.Name
            except Exception:
                pass

            result = SnapshotUndo.pop(app, doc_name)
            return json.dumps(result)
        except Exception as e:
            return json.dumps({"error": str(e)})

    elif sys.platform == "darwin":
        try:
            from word_document_server.core.word_mac import (
                find_document as mac_find_document,
                get_word_app as mac_get_word_app,
                _run_applescript,
                _escape_as,
            )
            from word_document_server.core.word_com import SnapshotUndo

            mac_get_word_app()
            doc_info = mac_find_document(None, filename)
            doc_name = doc_info.get("name", "")

            result = SnapshotUndo.pop_path(doc_name)
            if not result["success"]:
                return json.dumps(result)

            restored_path = result["restored"]
            _run_applescript(f"""
tell application "Microsoft Word"
    close document "{_escape_as(doc_name)}" saving no
    open "{_escape_as(restored_path)}"
end tell
""")
            return json.dumps(result)
        except Exception as e:
            return json.dumps({"error": str(e)})

    else:
        return json.dumps({"error": "Word automation is only supported on Windows and macOS"})


async def word_list_snapshots(filename: str = None) -> str:
    """List all available snapshots for a document (or all documents).

    Args:
        filename: Filter to snapshots for this document. None = list all.

    Returns:
        JSON array of snapshots with id, document name, and timestamp.
    """
    try:
        from word_document_server.core.word_com import SnapshotUndo

        doc_name = None
        if filename:
            doc_name = os.path.basename(filename)

        snapshots = SnapshotUndo.list(doc_name)
        return json.dumps({
            "success": True,
            "count": len(snapshots),
            "snapshots": snapshots,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_clear_snapshots(filename: str = None) -> str:
    """Delete all snapshots for a document (or all snapshots).

    Args:
        filename: Clear snapshots only for this document. None = clear everything.

    Returns:
        JSON confirming how many snapshots were removed.
    """
    try:
        from word_document_server.core.word_com import SnapshotUndo

        doc_name = None
        if filename:
            doc_name = os.path.basename(filename)

        before = len(SnapshotUndo.list(doc_name))
        SnapshotUndo.clear(doc_name)
        return json.dumps({
            "success": True,
            "cleared": before,
            "scope": doc_name or "all",
        })
    except Exception as e:
        return json.dumps({"error": str(e)})
