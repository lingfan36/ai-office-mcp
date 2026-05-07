"""Cross-platform file/session tools — openpyxl backend (no Office required)."""
import json
import os
from pathlib import Path
from excel_mcp.core.openpyxl_backend import OpenpyxlSession, OpenpyxlSnapshotUndo


def list_sessions() -> str:
    """List all open workbooks."""
    try:
        sessions = OpenpyxlSession.list_all()
        return json.dumps({"success": True, "sessions": sessions, "count": len(sessions)})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def open_workbook(path: str, read_only: bool = False, password: str = None) -> str:
    """Open a workbook by file path.

    Note: password-protected files are not supported in cross-platform mode.
    """
    try:
        if password:
            return json.dumps({
                "success": False,
                "error": "Password-protected files require Windows + Excel (COM mode).",
            })
        result = OpenpyxlSession.open(path, read_only=read_only)
        return json.dumps({"success": True, **result})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def close_workbook(workbook: str = None, save: bool = True) -> str:
    """Close a workbook, optionally saving changes."""
    try:
        name = OpenpyxlSession.close(workbook, save=save)
        OpenpyxlSnapshotUndo.clear(name)
        return json.dumps({"success": True, "closed": name, "saved": save})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def create_workbook(path: str, macro_enabled: bool = False) -> str:
    """Create a new empty workbook and save it to *path*.

    Note: macro-enabled (.xlsm) workbooks are not supported in cross-platform mode.
    """
    try:
        if macro_enabled:
            return json.dumps({
                "success": False,
                "error": "Macro-enabled workbooks require Windows + Excel (COM mode).",
            })
        result = OpenpyxlSession.create(path)
        return json.dumps({"success": True, **result})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def save_workbook(workbook: str = None, save_as: str = None) -> str:
    """Save the workbook, optionally to a new path (Save As)."""
    try:
        name, _ = OpenpyxlSession.get(workbook)
        target = OpenpyxlSession.save(workbook, save_as=save_as)
        return json.dumps({"success": True, "name": name, "path": target, "saved": True})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def get_workbook_info(workbook: str = None) -> str:
    """Return metadata about a workbook: sheets, named ranges, protection status."""
    try:
        name, session = OpenpyxlSession.get(workbook)
        wb = session["wb"]
        sheets = [
            {
                "name": ws.title,
                "visible": ws.sheet_state == "visible",
                "very_hidden": ws.sheet_state == "veryHidden",
                "tab_color": (f"#{ws.sheet_properties.tabColor.rgb[2:]}"
                              if ws.sheet_properties.tabColor else None),
            }
            for ws in wb.worksheets
        ]
        named = [
            {"name": nr, "ref": wb.defined_names[nr].attr_text}
            for nr in wb.defined_names
        ]
        info = {
            "name": name,
            "path": session["path"],
            "saved": True,
            "read_only": session["read_only"],
            "sheets": sheets,
            "named_ranges": named,
            "sheet_count": len(wb.sheetnames),
        }
        return json.dumps({"success": True, "info": info})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def export_pdf(output_path: str, sheet: str = None, workbook: str = None) -> str:
    """Export to PDF — requires Windows + Excel in COM mode."""
    return json.dumps({
        "success": False,
        "error": "PDF export requires Windows + Excel (COM mode). "
                 "On cross-platform mode, save as .xlsx and convert externally.",
    })


def protect_workbook(password: str = None, structure: bool = True,
                     windows: bool = False, workbook: str = None) -> str:
    """Workbook-level protection — requires Windows + Excel in COM mode."""
    return json.dumps({
        "success": False,
        "error": "Workbook protection requires Windows + Excel (COM mode).",
    })


def unprotect_workbook(password: str = None, workbook: str = None) -> str:
    """Remove workbook protection — requires Windows + Excel in COM mode."""
    return json.dumps({
        "success": False,
        "error": "Workbook protection requires Windows + Excel (COM mode).",
    })
