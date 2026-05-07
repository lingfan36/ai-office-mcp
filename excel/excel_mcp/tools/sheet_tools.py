"""Worksheet management tools (16 operations)."""
import json
from excel_mcp.core.excel_com import get_excel_app, find_workbook, find_sheet, SnapshotUndo
from excel_mcp.core.constants import xlSheetVisible, xlSheetHidden, xlSheetVeryHidden


def list_sheets(workbook: str = None) -> str:
    """List all worksheets in a workbook."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        sheets = []
        for ws in wb.Worksheets:
            sheets.append({
                "name": ws.Name,
                "index": ws.Index,
                "visible": ws.Visible == xlSheetVisible,
                "very_hidden": ws.Visible == xlSheetVeryHidden,
                "tab_color": (f"#{ws.Tab.Color & 0xFF:02X}{(ws.Tab.Color >> 8) & 0xFF:02X}{(ws.Tab.Color >> 16) & 0xFF:02X}"
                              if ws.Tab.ColorIndex != -4142 else None),
            })
        return json.dumps({"success": True, "sheets": sheets, "count": len(sheets)})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def create_sheet(name: str, position: int = None, workbook: str = None) -> str:
    """Create a new worksheet. position: 1-based index (None = last)."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        if position is not None:
            before = wb.Worksheets(position) if position <= wb.Worksheets.Count else None
            ws = wb.Worksheets.Add(Before=before) if before else wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        else:
            ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = name
        return json.dumps({"success": True, "name": ws.Name, "index": ws.Index})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def rename_sheet(old_name: str, new_name: str, workbook: str = None) -> str:
    """Rename a worksheet."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, old_name)
        ws.Name = new_name
        return json.dumps({"success": True, "old_name": old_name, "new_name": new_name})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def copy_sheet(sheet: str, new_name: str = None, position: int = None, workbook: str = None) -> str:
    """Copy a worksheet within the same workbook.

    Uses a COM-first path; falls back to openpyxl (save → copy → reopen) when the
    COM host (e.g. WPS) routes ws.Copy() to a new workbook instead of in-place.
    """
    import os, openpyxl as _xl
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        sheet_name = ws.Name          # capture before any close
        wb_path = wb.FullName
        after = wb.Worksheets(position) if position else wb.Worksheets(wb.Worksheets.Count)
        after_idx = after.Index
        wb_names_before = {wkb.Name for wkb in app.Workbooks}
        count_before = wb.Worksheets.Count

        ws.Copy(After=after)

        wb_names_after = {wkb.Name for wkb in app.Workbooks}
        orphan_names = wb_names_after - wb_names_before

        if orphan_names:
            # WPS opened a new workbook instead of copying in-place.
            # Suppress dialogs around all close/reopen ops to prevent blocking.
            try:
                app.DisplayAlerts = False
            except Exception:
                pass
            try:
                for name in list(orphan_names):
                    try:
                        app.Workbooks(name).Close(SaveChanges=False)
                    except Exception:
                        pass
                # Persist then close without prompt (we just saved)
                wb.Save()
                wb.Close(SaveChanges=False)
            finally:
                try:
                    app.DisplayAlerts = True
                except Exception:
                    pass

            # Use openpyxl to copy the sheet (avoids all WPS COM copy limitations)
            owb = _xl.load_workbook(wb_path)
            target_name = new_name or (sheet_name + " (2)")
            new_ows = owb.copy_worksheet(owb[sheet_name])
            new_ows.title = target_name
            owb.save(wb_path)

            try:
                app.DisplayAlerts = False
            except Exception:
                pass
            try:
                wb = app.Workbooks.Open(wb_path)
            finally:
                try:
                    app.DisplayAlerts = True
                except Exception:
                    pass
            new_ws = find_sheet(wb, target_name)
        elif wb.Worksheets.Count > count_before:
            try:
                new_ws = wb.Worksheets(after_idx + 1)
            except Exception:
                new_ws = wb.Worksheets(wb.Worksheets.Count)
            if new_name:
                new_ws.Name = new_name
        else:
            raise RuntimeError("copy_sheet: no new worksheet was created")

        return json.dumps({"success": True, "original": sheet_name, "copy": new_ws.Name, "index": new_ws.Index})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def move_sheet(sheet: str, position: int, workbook: str = None) -> str:
    """Move a worksheet to a new position (1-based)."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        if position >= wb.Worksheets.Count:
            ws.Move(After=wb.Worksheets(wb.Worksheets.Count))
        else:
            ws.Move(Before=wb.Worksheets(position))
        return json.dumps({"success": True, "sheet": sheet, "new_index": ws.Index})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def delete_sheet(sheet: str, workbook: str = None) -> str:
    """Delete a worksheet (irreversible — snapshot taken first)."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        app.DisplayAlerts = False
        ws.Delete()
        app.DisplayAlerts = True
        return json.dumps({"success": True, "deleted": sheet})
    except Exception as e:
        app.DisplayAlerts = True
        return json.dumps({"success": False, "error": str(e)})


def set_tab_color(sheet: str, color: str = None, workbook: str = None) -> str:
    """Set sheet tab color (#RRGGBB). Pass color=None to clear."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        if color is None:
            ws.Tab.ColorIndex = -4142  # xlColorIndexNone
        else:
            h = color.lstrip("#")
            r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
            ws.Tab.Color = r + g * 256 + b * 65536
        return json.dumps({"success": True, "sheet": sheet, "color": color})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_sheet_visibility(sheet: str, visibility: str = "visible", workbook: str = None) -> str:
    """Set sheet visibility. visibility: 'visible' | 'hidden' | 'very_hidden'."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        vis_map = {"visible": xlSheetVisible, "hidden": xlSheetHidden, "very_hidden": xlSheetVeryHidden}
        ws.Visible = vis_map.get(visibility, xlSheetVisible)
        return json.dumps({"success": True, "sheet": sheet, "visibility": visibility})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def copy_sheet_to_file(sheet: str, dest_path: str, dest_sheet_name: str = None,
                       workbook: str = None) -> str:
    """Copy a worksheet into another workbook file."""
    try:
        import os
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        abs_dest = os.path.abspath(dest_path)
        if os.path.exists(abs_dest):
            dest_wb = app.Workbooks.Open(abs_dest)
        else:
            dest_wb = app.Workbooks.Add()
            dest_wb.SaveAs(abs_dest)
        ws.Copy(After=dest_wb.Worksheets(dest_wb.Worksheets.Count))
        new_ws = dest_wb.Worksheets(dest_wb.Worksheets.Count)
        if dest_sheet_name:
            new_ws.Name = dest_sheet_name
        dest_wb.Save()
        return json.dumps({"success": True, "sheet": new_ws.Name, "dest": abs_dest})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def protect_sheet(sheet: str, password: str = None, workbook: str = None) -> str:
    """Protect a worksheet, optionally with a password."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        ws.Protect(Password=password or "")
        return json.dumps({"success": True, "sheet": sheet, "protected": True})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def unprotect_sheet(sheet: str, password: str = None, workbook: str = None) -> str:
    """Unprotect a worksheet."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        ws.Unprotect(Password=password or "")
        return json.dumps({"success": True, "sheet": sheet, "protected": False})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})
