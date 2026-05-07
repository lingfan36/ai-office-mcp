"""Excel Tables (ListObjects) tools (27 operations)."""
import json
from excel_mcp.core.excel_com import get_excel_app, find_workbook, find_sheet, SnapshotUndo


def _get_table(ws, table_name: str):
    for lo in ws.ListObjects:
        if lo.Name == table_name:
            return lo
    raise ValueError(f"Table not found: {table_name!r}")


def list_tables(sheet: str = None, workbook: str = None) -> str:
    """List all Excel Tables on a sheet (or all sheets if sheet=None)."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        results = []
        sheets = [find_sheet(wb, sheet)] if sheet else list(wb.Worksheets)
        for ws in sheets:
            for lo in ws.ListObjects:
                results.append({
                    "name": lo.Name,
                    "sheet": ws.Name,
                    "address": lo.Range.Address,
                    "rows": lo.ListRows.Count,
                    "cols": lo.ListColumns.Count,
                    "has_totals": lo.ShowTotals,
                })
        return json.dumps({"success": True, "tables": results, "count": len(results)})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def create_table(address: str, name: str, has_headers: bool = True,
                 style: str = "TableStyleMedium9", sheet: str = None, workbook: str = None) -> str:
    """Convert a range to an Excel Table."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        rng = ws.Range(address)
        lo = ws.ListObjects.Add(SourceType=1, Source=rng,
                                 XlListObjectHasHeaders=1 if has_headers else 2)
        lo.Name = name
        if style:
            lo.TableStyle = style
        return json.dumps({"success": True, "name": lo.Name, "address": lo.Range.Address})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def rename_table(table_name: str, new_name: str, sheet: str = None, workbook: str = None) -> str:
    """Rename an Excel Table."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        lo = _get_table(ws, table_name)
        lo.Name = new_name
        return json.dumps({"success": True, "old_name": table_name, "new_name": new_name})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def resize_table(table_name: str, new_address: str, sheet: str = None, workbook: str = None) -> str:
    """Resize an Excel Table to a new range address."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        lo = _get_table(ws, table_name)
        lo.Resize(ws.Range(new_address))
        return json.dumps({"success": True, "table": table_name, "new_address": lo.Range.Address})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def delete_table(table_name: str, delete_data: bool = False, sheet: str = None, workbook: str = None) -> str:
    """Remove an Excel Table (optionally deleting the underlying data)."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        lo = _get_table(ws, table_name)
        if delete_data:
            lo.Range.Delete()
        else:
            lo.Unlist()
        return json.dumps({"success": True, "deleted": table_name, "data_deleted": delete_data})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def apply_table_style(table_name: str, style: str, sheet: str = None, workbook: str = None) -> str:
    """Apply a built-in table style, e.g. 'TableStyleLight1' .. 'TableStyleDark11'."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        lo = _get_table(ws, table_name)
        lo.TableStyle = style
        return json.dumps({"success": True, "table": table_name, "style": style})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def toggle_totals_row(table_name: str, show: bool = True, sheet: str = None, workbook: str = None) -> str:
    """Show or hide the totals row of an Excel Table."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        lo = _get_table(ws, table_name)
        lo.ShowTotals = show
        return json.dumps({"success": True, "table": table_name, "totals_row": show})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_column_total(table_name: str, column_name: str, function: str = "sum",
                     sheet: str = None, workbook: str = None) -> str:
    """Set the aggregation function for a totals row cell.
    function: 'sum'|'count'|'average'|'max'|'min'|'none'
    """
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        lo = _get_table(ws, table_name)
        lo.ShowTotals = True
        func_map = {"sum": 9, "count": 2, "average": 1, "max": 4, "min": 5,
                    "stddev": 7, "var": 10, "none": 0}
        col = None
        for lc in lo.ListColumns:
            if lc.Name == column_name:
                col = lc
                break
        if col is None:
            raise ValueError(f"Column not found: {column_name!r}")
        col.TotalsCalculation = func_map.get(function, 0)
        return json.dumps({"success": True, "table": table_name, "column": column_name, "function": function})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def append_rows_to_table(table_name: str, data: list[list], sheet: str = None, workbook: str = None) -> str:
    """Append rows of data to the bottom of an Excel Table."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        lo = _get_table(ws, table_name)
        for row_data in data:
            new_row = lo.ListRows.Add()
            for i, val in enumerate(row_data, start=1):
                new_row.Range.Cells(1, i).Value = val
        return json.dumps({"success": True, "table": table_name, "rows_added": len(data)})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def get_table_data(table_name: str, visible_only: bool = False,
                   sheet: str = None, workbook: str = None) -> str:
    """Retrieve all data from an Excel Table as a 2-D list (including headers)."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        lo = _get_table(ws, table_name)
        headers = [lc.Name for lc in lo.ListColumns]
        rows = []
        for lr in lo.ListRows:
            if visible_only and lr.Range.EntireRow.Hidden:
                continue
            row = []
            for cell in lr.Range.Cells:
                v = cell.Value
                row.append(str(v) if hasattr(v, "year") else v)
            rows.append(row)
        return json.dumps({"success": True, "headers": headers, "rows": rows,
                           "total_rows": len(rows), "cols": len(headers)})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def apply_table_filter(table_name: str, column_name: str, criteria: str,
                       sheet: str = None, workbook: str = None) -> str:
    """Apply an AutoFilter criteria to a table column.
    criteria examples: '>100', '=Active', '<>""'
    """
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        lo = _get_table(ws, table_name)
        col_idx = None
        for i, lc in enumerate(lo.ListColumns, start=1):
            if lc.Name == column_name:
                col_idx = i
                break
        if col_idx is None:
            raise ValueError(f"Column not found: {column_name!r}")
        lo.Range.AutoFilter(Field=col_idx, Criteria1=criteria)
        return json.dumps({"success": True, "table": table_name, "column": column_name, "criteria": criteria})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def clear_table_filters(table_name: str, sheet: str = None, workbook: str = None) -> str:
    """Remove all AutoFilters from an Excel Table."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        lo = _get_table(ws, table_name)
        lo.AutoFilter.ShowAllData()
        return json.dumps({"success": True, "table": table_name, "filters_cleared": True})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def sort_table(table_name: str, column_name: str, order: str = "asc",
               sheet: str = None, workbook: str = None) -> str:
    """Sort an Excel Table by a column. order: 'asc' | 'desc'."""
    try:
        from excel_mcp.core.constants import xlAscending, xlDescending, xlGuess
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        lo = _get_table(ws, table_name)
        col = None
        for lc in lo.ListColumns:
            if lc.Name == column_name:
                col = lc
                break
        if col is None:
            raise ValueError(f"Column not found: {column_name!r}")
        sort_order = xlAscending if order == "asc" else xlDescending
        lo.Sort.SortFields.Clear()
        lo.Sort.SortFields.Add(Key=col.Range, SortOn=0, Order=sort_order)
        lo.Sort.Header = 1
        lo.Sort.Apply()
        return json.dumps({"success": True, "table": table_name, "column": column_name, "order": order})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def add_table_column(table_name: str, column_name: str, position: int = None,
                     formula: str = None, sheet: str = None, workbook: str = None) -> str:
    """Add a column to an Excel Table, optionally with a calculated formula."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        lo = _get_table(ws, table_name)
        if position:
            lc = lo.ListColumns.Add(position)
        else:
            lc = lo.ListColumns.Add()
        lc.Name = column_name
        if formula:
            lc.DataBodyRange.Formula = formula
        return json.dumps({"success": True, "table": table_name, "column": column_name, "index": lc.Index})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def delete_table_column(table_name: str, column_name: str,
                        sheet: str = None, workbook: str = None) -> str:
    """Delete a column from an Excel Table."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        lo = _get_table(ws, table_name)
        for lc in lo.ListColumns:
            if lc.Name == column_name:
                lc.Delete()
                return json.dumps({"success": True, "deleted": column_name})
        raise ValueError(f"Column not found: {column_name!r}")
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})
