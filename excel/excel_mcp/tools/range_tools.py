"""Range operations — values, formulas, formatting, validation, structure (46 ops)."""
import json
from typing import Any
from excel_mcp.core.excel_com import get_excel_app, find_workbook, find_sheet, SnapshotUndo
from excel_mcp.core.constants import (
    xlAscending, xlDescending, xlGuess, xlYes, xlNo,
    xlHAlignCenter, xlHAlignLeft, xlHAlignRight, xlHAlignGeneral,
    xlVAlignBottom, xlVAlignCenter, xlVAlignTop,
    xlContinuous, xlNone, xlThin, xlMedium, xlThick,
    xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight,
    xlInsideVertical, xlInsideHorizontal,
    xlValidateList, xlValidateWholeNumber, xlValidateDecimal,
)

# ── helpers ──────────────────────────────────────────────────────────────────

def _rng(ws, address: str):
    return ws.Range(address)

def _color(hex_str: str) -> int:
    """Convert #RRGGBB to Excel BGR int."""
    h = hex_str.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return r + g * 256 + b * 65536

HALIGN = {"center": xlHAlignCenter, "left": xlHAlignLeft, "right": xlHAlignRight, "general": xlHAlignGeneral}
VALIGN = {"bottom": xlVAlignBottom, "center": xlVAlignCenter, "top": xlVAlignTop}
BORDER_IDX = {
    "left": xlEdgeLeft, "top": xlEdgeTop, "bottom": xlEdgeBottom, "right": xlEdgeRight,
    "inner_v": xlInsideVertical, "inner_h": xlInsideHorizontal,
}

# ── values / formulas ─────────────────────────────────────────────────────────

def get_values(address: str, sheet: str = None, workbook: str = None) -> str:
    """Read cell values from a range. Returns 2-D list."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        rng = _rng(ws, address)
        raw = rng.Value
        if isinstance(raw, tuple):
            data = [list(row) for row in raw]
        else:
            data = [[raw]]
        # Convert COM dates/None to serialisable types
        cleaned = []
        for row in data:
            cleaned.append([str(v) if hasattr(v, "year") else v for v in row])
        return json.dumps({"success": True, "address": address, "values": cleaned,
                           "rows": rng.Rows.Count, "cols": rng.Columns.Count})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_values(address: str, values: list[list[Any]], sheet: str = None,
               workbook: str = None, snapshot: bool = True) -> str:
    """Write a 2-D list of values into a range."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        if snapshot:
            SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        rng = ws.Range(address)
        # Resize target if values matrix is larger
        rows, cols = len(values), max(len(r) for r in values)
        rng = ws.Range(rng.Cells(1, 1), rng.Cells(rows, cols))
        rng.Value = values
        return json.dumps({"success": True, "address": rng.Address, "rows": rows, "cols": cols})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def get_formulas(address: str, sheet: str = None, workbook: str = None) -> str:
    """Read cell formulas from a range."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        rng = _rng(ws, address)
        raw = rng.Formula
        data = [list(row) for row in raw] if isinstance(raw, tuple) else [[raw]]
        return json.dumps({"success": True, "address": address, "formulas": data})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_formulas(address: str, formulas: list[list[str]], sheet: str = None,
                 workbook: str = None, snapshot: bool = True) -> str:
    """Write formulas into a range. Each cell value should start with '='."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        if snapshot:
            SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        rows, cols = len(formulas), max(len(r) for r in formulas)
        rng = ws.Range(ws.Cells(ws.Range(address).Row, ws.Range(address).Column),
                       ws.Cells(ws.Range(address).Row + rows - 1, ws.Range(address).Column + cols - 1))
        rng.Formula = formulas
        return json.dumps({"success": True, "address": rng.Address})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def clear_range(address: str, clear_type: str = "all", sheet: str = None, workbook: str = None) -> str:
    """Clear a range. clear_type: 'all' | 'contents' | 'formats'."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        rng = _rng(ws, address)
        {"all": rng.Clear, "contents": rng.ClearContents, "formats": rng.ClearFormats}[clear_type]()
        return json.dumps({"success": True, "cleared": address, "type": clear_type})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def copy_range(src_address: str, dst_address: str, copy_type: str = "all",
               src_sheet: str = None, dst_sheet: str = None, workbook: str = None) -> str:
    """Copy a range. copy_type: 'all' | 'values' | 'formulas' | 'formats'."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        src_ws = find_sheet(wb, src_sheet)
        dst_ws = find_sheet(wb, dst_sheet) if dst_sheet else src_ws
        src = _rng(src_ws, src_address)
        dst = _rng(dst_ws, dst_address)
        if copy_type == "all":
            src.Copy(dst)
        elif copy_type == "values":
            dst.Value = src.Value
        elif copy_type == "formulas":
            dst.Formula = src.Formula
        elif copy_type == "formats":
            src.Copy()
            dst.PasteSpecial(Paste=-4122)  # xlPasteFormats
            app.CutCopyMode = False
        return json.dumps({"success": True, "from": src_address, "to": dst_address})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def insert_rows(row: int, count: int = 1, sheet: str = None, workbook: str = None) -> str:
    """Insert *count* rows before the given row number."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        ws.Rows(f"{row}:{row + count - 1}").Insert()
        return json.dumps({"success": True, "inserted_before": row, "count": count})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def delete_rows(row: int, count: int = 1, sheet: str = None, workbook: str = None) -> str:
    """Delete *count* rows starting at the given row number."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        ws.Rows(f"{row}:{row + count - 1}").Delete()
        return json.dumps({"success": True, "deleted_from": row, "count": count})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def insert_columns(col: int, count: int = 1, sheet: str = None, workbook: str = None) -> str:
    """Insert *count* columns before the given column number (1-based)."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        ws.Columns(f"{col}:{col + count - 1}").Insert()
        return json.dumps({"success": True, "inserted_before_col": col, "count": count})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def delete_columns(col: int, count: int = 1, sheet: str = None, workbook: str = None) -> str:
    """Delete *count* columns starting at the given column number (1-based)."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        ws.Columns(f"{col}:{col + count - 1}").Delete()
        return json.dumps({"success": True, "deleted_from_col": col, "count": count})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def find_replace(find_text: str, replace_text: str = None, address: str = None,
                 match_case: bool = False, match_entire_cell: bool = False,
                 sheet: str = None, workbook: str = None) -> str:
    """Find (and optionally replace) text in a range or whole sheet."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        scope = _rng(ws, address) if address else ws.UsedRange
        if replace_text is not None:
            SnapshotUndo.push(wb)
            scope.Replace(find_text, replace_text, MatchCase=match_case,
                          LookAt=2 if match_entire_cell else 1)
            return json.dumps({"success": True, "action": "replace", "find": find_text, "replace": replace_text})
        else:
            cell = scope.Find(find_text, MatchCase=match_case, LookAt=2 if match_entire_cell else 1)
            if cell:
                return json.dumps({"success": True, "found": True, "address": cell.Address})
            return json.dumps({"success": True, "found": False})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def sort_range(address: str, key_col: int = 1, order: str = "asc", has_header: bool = True,
               sheet: str = None, workbook: str = None) -> str:
    """Sort a range by a column. order: 'asc' | 'desc'."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        rng = _rng(ws, address)
        sort_order = xlAscending if order == "asc" else xlDescending
        if has_header:
            # WPS ignores Header=xlYes AND expands the sort range to the
            # CurrentRegion, sorting the header row along with the data.
            # Save the header row values, sort the full range (Header=xlNo so
            # the sort key is the original row-2 cell), then write the saved
            # values back to row 1.
            num_cols = rng.Columns.Count
            header_vals = [rng.Cells(1, c).Value for c in range(1, num_cols + 1)]
            rng.Sort(Key1=rng.Cells(2, key_col), Order1=sort_order, Header=xlNo)
            for c, val in enumerate(header_vals, 1):
                rng.Cells(1, c).Value = val
        else:
            rng.Sort(Key1=rng.Cells(1, key_col), Order1=sort_order, Header=xlNo)
        return json.dumps({"success": True, "sorted": address, "key_col": key_col, "order": order})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def get_used_range(sheet: str = None, workbook: str = None) -> str:
    """Return address and dimensions of the used range on a sheet."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        ur = ws.UsedRange
        return json.dumps({"success": True, "address": ur.Address,
                           "rows": ur.Rows.Count, "cols": ur.Columns.Count,
                           "first_row": ur.Row, "first_col": ur.Column})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def get_current_region(address: str, sheet: str = None, workbook: str = None) -> str:
    """Return the contiguous data block (CurrentRegion) around a cell."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        cr = _rng(ws, address).CurrentRegion
        return json.dumps({"success": True, "address": cr.Address,
                           "rows": cr.Rows.Count, "cols": cr.Columns.Count})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def get_range_info(address: str, sheet: str = None, workbook: str = None) -> str:
    """Return address, dimensions, and basic metadata about a range."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        rng = _rng(ws, address)
        return json.dumps({"success": True, "address": rng.Address,
                           "rows": rng.Rows.Count, "cols": rng.Columns.Count,
                           "first_row": rng.Row, "first_col": rng.Column,
                           "count": rng.Count})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def get_number_format(address: str, sheet: str = None, workbook: str = None) -> str:
    """Get number format string of a range."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        rng = _rng(ws, address)
        return json.dumps({"success": True, "address": address, "number_format": rng.NumberFormat})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_number_format(address: str, number_format: str, sheet: str = None, workbook: str = None) -> str:
    """Set number format of a range, e.g. '#,##0.00', 'yyyy-mm-dd', '0%'."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        _rng(ws, address).NumberFormat = number_format
        return json.dumps({"success": True, "address": address, "number_format": number_format})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def format_range(address: str, sheet: str = None, workbook: str = None,
                 bold: bool = None, italic: bool = None, underline: bool = None,
                 font_size: float = None, font_name: str = None,
                 font_color: str = None, bg_color: str = None,
                 h_align: str = None, v_align: str = None,
                 wrap_text: bool = None, number_format: str = None,
                 border: str = None, border_weight: str = "thin") -> str:
    """Apply font, fill, alignment, and border formatting to a range.

    border: 'all' | 'outer' | 'inner' | 'none'
    border_weight: 'thin' | 'medium' | 'thick'
    Colors as '#RRGGBB'.
    """
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        rng = _rng(ws, address)

        if bold is not None:
            rng.Font.Bold = bold
        if italic is not None:
            rng.Font.Italic = italic
        if underline is not None:
            rng.Font.Underline = 2 if underline else -4142
        if font_size is not None:
            rng.Font.Size = font_size
        if font_name is not None:
            rng.Font.Name = font_name
        if font_color is not None:
            rng.Font.Color = _color(font_color)
        if bg_color is not None:
            rng.Interior.Color = _color(bg_color)
        if h_align is not None:
            rng.HorizontalAlignment = HALIGN.get(h_align, xlHAlignGeneral)
        if v_align is not None:
            rng.VerticalAlignment = VALIGN.get(v_align, xlVAlignBottom)
        if wrap_text is not None:
            rng.WrapText = wrap_text
        if number_format is not None:
            rng.NumberFormat = number_format

        if border:
            weight_map = {"thin": xlThin, "medium": xlMedium, "thick": xlThick}
            w = weight_map.get(border_weight, xlThin)
            edges = (list(BORDER_IDX.values())
                     if border == "all"
                     else [xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight]
                     if border == "outer"
                     else [xlInsideVertical, xlInsideHorizontal]
                     if border == "inner"
                     else [])
            for idx in edges:
                b = rng.Borders(idx)
                if border == "none":
                    b.LineStyle = xlNone
                else:
                    b.LineStyle = xlContinuous
                    b.Weight = w

        return json.dumps({"success": True, "formatted": address})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def add_validation(address: str, validation_type: str, formula1: str,
                   formula2: str = None, operator: str = "between",
                   error_message: str = None,
                   prompt: str = None, sheet: str = None, workbook: str = None) -> str:
    """Add data validation to a range.
    validation_type: 'list' | 'whole_number' | 'decimal' | 'custom'
    operator (for numeric types): 'between'|'not_between'|'equal'|'not_equal'|
                                  'greater'|'less'|'greater_equal'|'less_equal'
    """
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        rng = _rng(ws, address)
        rng.Validation.Delete()
        type_map = {"list": xlValidateList, "whole_number": xlValidateWholeNumber,
                    "decimal": xlValidateDecimal, "custom": 7}
        vtype = type_map.get(validation_type, xlValidateList)
        op_map = {"between": 1, "not_between": 2, "equal": 3, "not_equal": 4,
                  "greater": 5, "less": 6, "greater_equal": 7, "less_equal": 8}
        op_val = op_map.get(operator, 1)

        # WPS requires Operator in Validation.Add() for all types (including list).
        kwargs = {"Type": vtype, "Operator": op_val if validation_type != "list" else 1,
                  "Formula1": formula1}
        if formula2:
            kwargs["Formula2"] = formula2

        last_err = None
        # Try multiple call signatures — WPS COM does not accept AlertStyle in Add().
        for extra in [{}, {"AlertStyle": 1}]:
            try:
                rng.Validation.Add(**{**kwargs, **extra})
                last_err = None
                break
            except Exception as _e:
                last_err = _e
        if last_err is not None:
            raise last_err
        try:
            rng.Validation.ErrorStyle = 1  # xlValidAlertStop
        except Exception:
            pass
        if error_message:
            rng.Validation.ErrorMessage = error_message
            rng.Validation.ShowError = True
        if prompt:
            rng.Validation.InputMessage = prompt
            rng.Validation.ShowInput = True
        return json.dumps({"success": True, "address": address, "type": validation_type})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def remove_validation(address: str, sheet: str = None, workbook: str = None) -> str:
    """Remove data validation from a range."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        _rng(ws, address).Validation.Delete()
        return json.dumps({"success": True, "address": address})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def merge_cells(address: str, across: bool = False, sheet: str = None, workbook: str = None) -> str:
    """Merge cells in a range. across=True merges each row independently."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        _rng(ws, address).Merge(Across=across)
        return json.dumps({"success": True, "merged": address})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def unmerge_cells(address: str, sheet: str = None, workbook: str = None) -> str:
    """Unmerge cells in a range."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        _rng(ws, address).UnMerge()
        return json.dumps({"success": True, "unmerged": address})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def autofit(address: str, target: str = "both", sheet: str = None, workbook: str = None) -> str:
    """Auto-fit column widths and/or row heights. target: 'columns' | 'rows' | 'both'."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        rng = _rng(ws, address)
        if target in ("columns", "both"):
            rng.Columns.AutoFit()
        if target in ("rows", "both"):
            rng.Rows.AutoFit()
        return json.dumps({"success": True, "address": address, "target": target})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def add_hyperlink(address: str, url: str, display_text: str = None,
                  sheet: str = None, workbook: str = None) -> str:
    """Add a hyperlink to a cell."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        rng = _rng(ws, address)
        ws.Hyperlinks.Add(Anchor=rng, Address=url,
                          TextToDisplay=display_text or rng.Value or url)
        return json.dumps({"success": True, "address": address, "url": url})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def remove_hyperlinks(address: str, sheet: str = None, workbook: str = None) -> str:
    """Remove all hyperlinks from a range."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        _rng(ws, address).Hyperlinks.Delete()
        return json.dumps({"success": True, "address": address})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_cell_lock(address: str, locked: bool = True, sheet: str = None, workbook: str = None) -> str:
    """Set the locked state of cells (effective when sheet is protected)."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        _rng(ws, address).Locked = locked
        return json.dumps({"success": True, "address": address, "locked": locked})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


# ── Row / column sizing & visibility ─────────────────────────────────────────

def _col_letter(n: int) -> str:
    """Convert 1-based column number to Excel letter notation (A, B, … Z, AA, …)."""
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def set_row_height(row: int, height: float, count: int = 1,
                   sheet: str = None, workbook: str = None) -> str:
    """Set row height in points for *count* rows starting at *row*."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        ws.Rows(f"{row}:{row + count - 1}").RowHeight = height
        return json.dumps({"success": True, "row": row, "count": count, "height": height})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_column_width(col: int, width: float, count: int = 1,
                     sheet: str = None, workbook: str = None) -> str:
    """Set column width in character units for *count* columns starting at *col* (1-based)."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        # WPS requires letter-based column address ("A:C"), not numeric ("1:3")
        col_range = f"{_col_letter(col)}:{_col_letter(col + count - 1)}"
        ws.Columns(col_range).ColumnWidth = width
        return json.dumps({"success": True, "col": col, "count": count, "width": width})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def hide_rows(row: int, count: int = 1, sheet: str = None, workbook: str = None) -> str:
    """Hide *count* rows starting at *row*."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        ws.Rows(f"{row}:{row + count - 1}").Hidden = True
        return json.dumps({"success": True, "row": row, "count": count, "hidden": True})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def show_rows(row: int, count: int = 1, sheet: str = None, workbook: str = None) -> str:
    """Unhide *count* rows starting at *row*."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        ws.Rows(f"{row}:{row + count - 1}").Hidden = False
        return json.dumps({"success": True, "row": row, "count": count, "hidden": False})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def hide_columns(col: int, count: int = 1, sheet: str = None, workbook: str = None) -> str:
    """Hide *count* columns starting at *col* (1-based)."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        col_range = f"{_col_letter(col)}:{_col_letter(col + count - 1)}"
        ws.Columns(col_range).Hidden = True
        return json.dumps({"success": True, "col": col, "count": count, "hidden": True})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def show_columns(col: int, count: int = 1, sheet: str = None, workbook: str = None) -> str:
    """Unhide *count* columns starting at *col* (1-based)."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        col_range = f"{_col_letter(col)}:{_col_letter(col + count - 1)}"
        ws.Columns(col_range).Hidden = False
        return json.dumps({"success": True, "col": col, "count": count, "hidden": False})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


# ── Freeze panes ─────────────────────────────────────────────────────────────

def freeze_panes(row: int = 0, col: int = 0,
                 sheet: str = None, workbook: str = None) -> str:
    """Freeze rows above *row* and columns left of *col*.
    row=1, col=0 freezes only the top row; row=0, col=1 freezes only column A.
    row=0, col=0 is equivalent to unfreeze_panes.

    WPS COM blocks setting win.FreezePanes = True directly.  We work around
    this via the Excel-4 macro FREEZE.PANES, falling back to the COM property
    if that also fails.
    """
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        ws.Activate()
        win = app.ActiveWindow

        # Always unfreeze first (guard: skip if already unfrozen)
        try:
            if win.FreezePanes:
                win.FreezePanes = False
        except Exception:
            pass
        # Also clear via XL4 macro for WPS compatibility
        try:
            app.ExecuteExcel4Macro("FREEZE.PANES(FALSE)")
        except Exception:
            pass

        if row > 0 or col > 0:
            r = row + 1 if row > 0 else 1
            c = col + 1 if col > 0 else 1
            # Navigate to the freeze-point cell
            try:
                ws.Cells(r, c).Select()
            except Exception:
                try:
                    app.Goto(ws.Cells(r, c))
                except Exception:
                    pass
            # Primary: Excel-4 macro (works in WPS where COM property fails)
            freeze_ok = False
            try:
                app.ExecuteExcel4Macro("FREEZE.PANES(TRUE)")
                freeze_ok = True
            except Exception:
                pass
            # Fallback: COM property
            if not freeze_ok:
                win.FreezePanes = True

        return json.dumps({"success": True, "freeze_row": row, "freeze_col": col})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def unfreeze_panes(sheet: str = None, workbook: str = None) -> str:
    """Remove all freeze panes from a sheet."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        ws.Activate()
        win = app.ActiveWindow
        try:
            if win.FreezePanes:
                win.FreezePanes = False
        except Exception:
            pass
        try:
            app.ExecuteExcel4Macro("FREEZE.PANES(FALSE)")
        except Exception:
            pass
        return json.dumps({"success": True})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


# ── AutoFilter (range-level, non-Table) ──────────────────────────────────────

def apply_range_autofilter(address: str, field: int = None, criteria: str = None,
                            sheet: str = None, workbook: str = None) -> str:
    """Enable AutoFilter on a range and optionally filter by a column.

    *field* is 1-based column index within *address*.
    *criteria* is a string like '>100', 'Apple', or '=B2'.
    Call with only *address* to just enable the AutoFilter dropdowns.
    """
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        rng = _rng(ws, address)
        if field is not None and criteria is not None:
            rng.AutoFilter(Field=field, Criteria1=criteria)
        elif field is not None:
            rng.AutoFilter(Field=field)
        else:
            rng.AutoFilter()
        return json.dumps({"success": True, "address": address, "field": field, "criteria": criteria})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def clear_range_autofilter(sheet: str = None, workbook: str = None) -> str:
    """Clear all AutoFilter criteria on the sheet (show all rows) without removing the dropdowns."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        if ws.AutoFilterMode:
            ws.AutoFilter.ShowAllData()
        return json.dumps({"success": True})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def toggle_autofilter(sheet: str = None, workbook: str = None) -> str:
    """Toggle AutoFilter on/off for the sheet.

    WPS does not support setting AutoFilterMode = True directly; to turn the
    filter ON we call AutoFilter() on the UsedRange instead.
    """
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        if ws.AutoFilterMode:
            ws.AutoFilterMode = False
            new_state = False
        else:
            ws.UsedRange.AutoFilter()
            new_state = bool(ws.AutoFilterMode)
        return json.dumps({"success": True, "autofilter_on": new_state})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


# ── Group / ungroup rows and columns ─────────────────────────────────────────

def group_rows(row: int, count: int = 1, sheet: str = None, workbook: str = None) -> str:
    """Group *count* rows starting at *row* (creates outline/collapse group)."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        ws.Rows(f"{row}:{row + count - 1}").Group()
        return json.dumps({"success": True, "row": row, "count": count})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def ungroup_rows(row: int, count: int = 1, sheet: str = None, workbook: str = None) -> str:
    """Ungroup *count* rows starting at *row*."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        ws.Rows(f"{row}:{row + count - 1}").Ungroup()
        return json.dumps({"success": True, "row": row, "count": count})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def group_columns(col: int, count: int = 1, sheet: str = None, workbook: str = None) -> str:
    """Group *count* columns starting at *col* (1-based)."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        col_range = f"{_col_letter(col)}:{_col_letter(col + count - 1)}"
        ws.Columns(col_range).Group()
        return json.dumps({"success": True, "col": col, "count": count})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def ungroup_columns(col: int, count: int = 1, sheet: str = None, workbook: str = None) -> str:
    """Ungroup *count* columns starting at *col* (1-based)."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        col_range = f"{_col_letter(col)}:{_col_letter(col + count - 1)}"
        ws.Columns(col_range).Ungroup()
        return json.dumps({"success": True, "col": col, "count": count})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


# ── Remove duplicates ─────────────────────────────────────────────────────────

def remove_duplicates(address: str, columns: list = None, has_header: bool = True,
                      sheet: str = None, workbook: str = None) -> str:
    """Remove duplicate rows from a range.

    *columns* is a 1-based list of column indices (within *address*) to check
    for duplicates. Omit to check all columns.
    """
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        rng = _rng(ws, address)
        rows_before = rng.Rows.Count
        kwargs = {"Header": xlYes if has_header else xlNo}
        if columns:
            kwargs["Columns"] = columns
        rng.RemoveDuplicates(**kwargs)
        rows_after = ws.Range(address.split(":")[0]).CurrentRegion.Rows.Count
        removed = rows_before - rows_after
        return json.dumps({"success": True, "address": address, "rows_removed": removed})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})

