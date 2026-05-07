"""Snapshot-based undo tools — openpyxl backend."""
import json
from excel_mcp.core.openpyxl_backend import OpenpyxlSnapshotUndo


def undo_last(workbook: str = None) -> str:
    """Restore the workbook to its state before the last AI operation."""
    try:
        result = OpenpyxlSnapshotUndo.pop(workbook)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def list_undo_snapshots(workbook: str = None) -> str:
    """List available snapshots for undo."""
    try:
        snapshots = OpenpyxlSnapshotUndo.list(workbook)
        return json.dumps({"success": True, "snapshots": snapshots, "count": len(snapshots)})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def clear_undo_history(workbook: str = None) -> str:
    """Clear all snapshots for a workbook (or all if workbook=None)."""
    try:
        OpenpyxlSnapshotUndo.clear(workbook)
        return json.dumps({"success": True, "cleared": workbook or "all"})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})
