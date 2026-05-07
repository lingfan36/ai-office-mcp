"""Workbook / session management tools (6 operations)."""
import json
import os
from excel_mcp.core.excel_com import get_excel_app, find_workbook, SnapshotUndo
from excel_mcp.core.constants import xlOpenXMLWorkbook, xlOpenXMLWorkbookMacroEnabled


def list_sessions() -> str:
    """List all open workbooks (active sessions)."""
    try:
        app = get_excel_app()
        sessions = []
        for wb in app.Workbooks:
            sessions.append({
                "name": wb.Name,
                "path": wb.FullName,
                "saved": wb.Saved,
                "sheets": wb.Worksheets.Count,
                "active": wb.Name == app.ActiveWorkbook.Name if app.ActiveWorkbook else False,
            })
        return json.dumps({"success": True, "sessions": sessions, "count": len(sessions)})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def open_workbook(path: str, read_only: bool = False, password: str = None) -> str:
    """Open a workbook by file path. Returns workbook name."""
    try:
        app = get_excel_app()
        abs_path = os.path.abspath(path)

        # Close any workbooks WPS may have auto-recovered from a previous forced
        # exit.  If left open they become the ActiveWorkbook and break all
        # subsequent find_workbook(None) calls that expect our target file.
        try:
            app.DisplayAlerts = False
            for _wb in list(app.Workbooks):
                if _wb.FullName.lower() != abs_path.lower():
                    try:
                        _wb.Close(SaveChanges=False)
                    except Exception:
                        pass
        except Exception:
            pass
        finally:
            try:
                app.DisplayAlerts = True
            except Exception:
                pass

        kwargs = {"ReadOnly": read_only}
        if password:
            kwargs["Password"] = password
        # Temporarily suppress alerts only for the Open call (WPS format dialogs),
        # then restore so validation/pivot ops aren't affected.
        try:
            app.DisplayAlerts = False
        except Exception:
            pass
        try:
            wb = app.Workbooks.Open(abs_path, **kwargs)
        finally:
            try:
                app.DisplayAlerts = True
            except Exception:
                pass
        wb.Activate()  # ensure this workbook is the active one even if others are open
        return json.dumps({"success": True, "name": wb.Name, "path": wb.FullName, "sheets": wb.Worksheets.Count})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def close_workbook(workbook: str = None, save: bool = True) -> str:
    """Close a workbook, optionally saving changes."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        name = wb.Name
        wb.Close(SaveChanges=save)
        SnapshotUndo.clear(name)
        return json.dumps({"success": True, "closed": name, "saved": save})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def create_workbook(path: str, macro_enabled: bool = False) -> str:
    """Create a new empty workbook and save it to *path*."""
    try:
        app = get_excel_app()
        wb = app.Workbooks.Add()
        fmt = xlOpenXMLWorkbookMacroEnabled if macro_enabled else xlOpenXMLWorkbook
        abs_path = os.path.abspath(path)
        wb.SaveAs(abs_path, FileFormat=fmt)
        return json.dumps({"success": True, "name": wb.Name, "path": wb.FullName})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def save_workbook(workbook: str = None, save_as: str = None) -> str:
    """Save a workbook. Optionally save to a new path (Save As).

    When no explicit *save_as* path is given we use SaveAs(current_path) with
    DisplayAlerts suppressed.  This is more reliable than wb.Save() under WPS
    COM, which silently drops newly-created sheets (e.g. those containing pivot
    tables) when serialising to xlsx.
    """
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        target = os.path.abspath(save_as) if save_as else wb.FullName
        # Determine FileFormat from extension
        ext = os.path.splitext(target)[1].lower()
        fmt_map = {".xlsx": 51, ".xlsm": 52, ".xls": -4143, ".xlsb": 50}
        file_format = fmt_map.get(ext, 51)
        try:
            app.DisplayAlerts = False
        except Exception:
            pass
        try:
            wb.SaveAs(target, FileFormat=file_format)
        finally:
            try:
                app.DisplayAlerts = True
            except Exception:
                pass
        return json.dumps({"success": True, "name": wb.Name, "path": wb.FullName, "saved": True})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def get_workbook_info(workbook: str = None) -> str:
    """Return metadata about a workbook: sheets, named ranges, protection status."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        sheets = [{"name": ws.Name, "visible": ws.Visible == -1} for ws in wb.Worksheets]
        named = [{"name": n.Name, "ref": n.RefersTo} for n in wb.Names]
        info = {
            "name": wb.Name,
            "path": wb.FullName,
            "saved": wb.Saved,
            "read_only": wb.ReadOnly,
            "sheets": sheets,
            "named_ranges": named,
            "sheet_count": wb.Worksheets.Count,
        }
        return json.dumps({"success": True, "info": info})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def protect_workbook(password: str = None, structure: bool = True,
                     windows: bool = False, workbook: str = None) -> str:
    """Protect a workbook's structure and/or windows.

    *structure=True* prevents inserting, deleting, moving, or hiding sheets.
    *windows=True* prevents moving or resizing workbook windows.
    """
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        kwargs = {"Structure": structure, "Windows": windows}
        if password:
            kwargs["Password"] = password
        wb.Protect(**kwargs)
        return json.dumps({"success": True, "name": wb.Name,
                           "structure": structure, "windows": windows})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def unprotect_workbook(password: str = None, workbook: str = None) -> str:
    """Remove workbook-level protection."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        if password:
            wb.Unprotect(Password=password)
        else:
            wb.Unprotect()
        return json.dumps({"success": True, "name": wb.Name})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def export_pdf(output_path: str, sheet: str = None, workbook: str = None) -> str:
    """Export a workbook or a single sheet to PDF.

    If *sheet* is given, only that sheet is exported; otherwise the entire
    workbook is exported. *output_path* should end with '.pdf'.
    """
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        abs_path = os.path.abspath(output_path)
        xlTypePDF = 0
        if sheet:
            from excel_mcp.core.excel_com import find_sheet
            ws = find_sheet(wb, sheet)
            ws.ExportAsFixedFormat(Type=xlTypePDF, Filename=abs_path)
        else:
            wb.ExportAsFixedFormat(Type=xlTypePDF, Filename=abs_path)
        return json.dumps({"success": True, "output": abs_path})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})
