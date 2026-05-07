"""PivotTable tools (30 operations)."""
import json
import re
from excel_mcp.core.excel_com import get_excel_app, find_workbook, find_sheet, SnapshotUndo
from excel_mcp.core.constants import (
    xlRowField, xlColumnField, xlPageField, xlDataField, xlHidden,
    xlAscending, xlDescending, AGGREGATION,
)


def _get_pivot(ws, pivot_name: str):
    for pt in ws.PivotTables():
        if pt.Name == pivot_name:
            return pt
    raise ValueError(f"PivotTable not found: {pivot_name!r}")


def _parse_r1c1(addr: str):
    """Parse 'R1C1:R17C7' → (start_row, start_col, end_row, end_col) or None."""
    m = re.match(r'R(\d+)C(\d+):R(\d+)C(\d+)', addr.strip().upper())
    if m:
        return int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
    return None


def _find_pivot_field(pt, wb, field_name: str):
    """Find a PivotField by name with WPS-specific fallbacks.

    WPS sometimes names fields using the first data row's cell values instead
    of the column header text.  When the standard name lookup fails we fall
    back to reading the source range's actual header row and matching by
    column index.
    """
    # 1. Standard name lookup
    try:
        return pt.PivotFields(field_name)
    except Exception:
        pass

    # 2. Case-insensitive / whitespace-tolerant scan
    try:
        for i in range(1, pt.PivotFields().Count + 1):
            f = pt.PivotFields(i)
            if f.Name.strip() == field_name.strip():
                return f
    except Exception:
        pass

    # 3. WPS fallback: read actual header row from SourceData range.
    # WPS stores SourceData in R1C1 notation (e.g. 'Sheet1!R1C1:R17C7').
    # ws.Range("R1C1:...") fails in WPS COM, so we parse manually and use
    # ws.Cells(row, col) to read the header cells instead.
    try:
        source = str(pt.SourceData)
        m = re.match(r"'?([^'!]+)'?!(.+)", source.strip())
        if m:
            sht_name = m.group(1).strip()
            addr = m.group(2).strip()
            for ws_cand in wb.Worksheets:
                if ws_cand.Name.strip() == sht_name.strip():
                    coords = _parse_r1c1(addr)
                    if coords:
                        r1, c1, r2, c2 = coords
                        num_cols = c2 - c1 + 1
                        for col_idx in range(1, num_cols + 1):
                            sheet_col = c1 + col_idx - 1
                            val = str(ws_cand.Cells(r1, sheet_col).Value or "").strip()
                            if val == field_name.strip():
                                return pt.PivotFields(col_idx)
                    else:
                        # A1 notation fallback (no $ signs)
                        clean_addr = addr.replace("$", "")
                        header_rng = ws_cand.Range(clean_addr).Rows(1)
                        for col_idx in range(1, header_rng.Columns.Count + 1):
                            val = str(header_rng.Cells(1, col_idx).Value or "").strip()
                            if val == field_name.strip():
                                return pt.PivotFields(col_idx)
                    break
    except Exception:
        pass

    return None


def list_pivot_tables(sheet: str = None, workbook: str = None) -> str:
    """List all PivotTables on a sheet or the entire workbook."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        sheets = [find_sheet(wb, sheet)] if sheet else list(wb.Worksheets)
        results = []
        for ws in sheets:
            for pt in ws.PivotTables():
                results.append({
                    "name": pt.Name,
                    "sheet": ws.Name,
                    "address": pt.TableRange2.Address,
                    "source": str(pt.SourceData),
                })
        return json.dumps({"success": True, "pivot_tables": results, "count": len(results)})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def create_pivot_table(source_address: str, dest_cell: str, name: str,
                       source_sheet: str = None, dest_sheet: str = None,
                       workbook: str = None) -> str:
    """Create a PivotTable from a range. dest_cell is the top-left of the output."""
    wb = None
    app = None
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        src_ws = find_sheet(wb, source_sheet)
        dst_ws = find_sheet(wb, dest_sheet) if dest_sheet else src_ws
        src_sheet_name = src_ws.Name
        # Use string-based SourceData for better WPS compatibility
        source_data_str = f"'{src_sheet_name}'!{source_address}"
        try:
            pc = wb.PivotCaches().Create(SourceType=1, SourceData=source_data_str)
        except Exception:
            # Fallback: Range object
            pc = wb.PivotCaches().Create(SourceType=1, SourceData=src_ws.Range(source_address))
        # Suppress the "overwrite existing data?" dialog that WPS shows when the
        # destination overlaps cells that already have content.  We only suppress
        # for this single call so that PivotCaches().Create() (which requires
        # DisplayAlerts=True in WPS) is unaffected.
        try:
            app.DisplayAlerts = False
        except Exception:
            pass
        try:
            pt = pc.CreatePivotTable(TableDestination=dst_ws.Range(dest_cell), TableName=name)
        finally:
            try:
                app.DisplayAlerts = True
            except Exception:
                pass
        return json.dumps({"success": True, "name": pt.Name, "sheet": dst_ws.Name,
                           "address": pt.TableRange2.Address})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})
    finally:
        # Always keep focus on the correct workbook even if pivot creation failed.
        # A failed CreatePivotTable call can cause WPS to switch the active workbook,
        # which would break all subsequent find_workbook(None) calls.
        if wb is not None:
            try:
                wb.Activate()
            except Exception:
                pass


def create_pivot_from_table(table_name: str, dest_cell: str, name: str,
                            table_sheet: str = None, dest_sheet: str = None,
                            workbook: str = None) -> str:
    """Create a PivotTable from an Excel Table (ListObject)."""
    wb = None
    app = None
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        src_ws = find_sheet(wb, table_sheet)
        lo = None
        for obj in src_ws.ListObjects:
            if obj.Name == table_name:
                lo = obj
                break
        if lo is None:
            raise ValueError(f"Table not found: {table_name!r}")
        dst_ws = find_sheet(wb, dest_sheet) if dest_sheet else src_ws
        pc = wb.PivotCaches().Create(SourceType=1, SourceData=lo.Range)
        try:
            app.DisplayAlerts = False
        except Exception:
            pass
        try:
            pt = pc.CreatePivotTable(TableDestination=dst_ws.Range(dest_cell), TableName=name)
        finally:
            try:
                app.DisplayAlerts = True
            except Exception:
                pass
        return json.dumps({"success": True, "name": pt.Name, "sheet": dst_ws.Name})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})
    finally:
        if wb is not None:
            try:
                wb.Activate()
            except Exception:
                pass


def add_pivot_field(pivot_name: str, field_name: str, area: str,
                    sheet: str = None, workbook: str = None) -> str:
    """Add a field to a PivotTable area.
    area: 'row' | 'column' | 'filter' | 'value'
    """
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        pt = _get_pivot(ws, pivot_name)
        area_map = {"row": xlRowField, "column": xlColumnField, "filter": xlPageField, "value": xlDataField}
        pf = _find_pivot_field(pt, wb, field_name)
        if pf is None:
            raise ValueError(f"Pivot field not found: {field_name!r}")
        pf.Orientation = area_map.get(area, xlRowField)
        return json.dumps({"success": True, "pivot": pivot_name, "field": field_name, "area": area})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def remove_pivot_field(pivot_name: str, field_name: str, sheet: str = None, workbook: str = None) -> str:
    """Remove a field from a PivotTable."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        pt = _get_pivot(ws, pivot_name)
        pf = _find_pivot_field(pt, wb, field_name)
        if pf is None:
            raise ValueError(f"Pivot field not found: {field_name!r}")
        pf.Orientation = xlHidden
        return json.dumps({"success": True, "pivot": pivot_name, "removed": field_name})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_pivot_aggregation(pivot_name: str, field_name: str, function: str = "sum",
                          sheet: str = None, workbook: str = None) -> str:
    """Set the aggregation function for a value field.
    function: 'sum'|'count'|'average'|'max'|'min'|'product'|'std_dev'|'var'
    """
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        pt = _get_pivot(ws, pivot_name)
        pf = _find_pivot_field(pt, wb, field_name)
        if pf is None:
            raise ValueError(f"Pivot field not found: {field_name!r}")
        pf.Function = AGGREGATION.get(function, AGGREGATION["sum"])
        return json.dumps({"success": True, "pivot": pivot_name, "field": field_name, "function": function})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_pivot_field_name(pivot_name: str, field_name: str, display_name: str,
                         sheet: str = None, workbook: str = None) -> str:
    """Rename a data field as displayed in the PivotTable."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        pt = _get_pivot(ws, pivot_name)
        pt.PivotFields(field_name).Caption = display_name
        return json.dumps({"success": True, "pivot": pivot_name, "field": field_name, "display": display_name})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_pivot_number_format(pivot_name: str, field_name: str, number_format: str,
                            sheet: str = None, workbook: str = None) -> str:
    """Apply a number format to a PivotTable value field."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        pt = _get_pivot(ws, pivot_name)
        pt.PivotFields(field_name).NumberFormat = number_format
        return json.dumps({"success": True, "pivot": pivot_name, "field": field_name, "format": number_format})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def filter_pivot_field(pivot_name: str, field_name: str, values: list[str],
                       sheet: str = None, workbook: str = None) -> str:
    """Filter a PivotTable row/column field to show only specified values."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        pt = _get_pivot(ws, pivot_name)
        pf = pt.PivotFields(field_name)
        values_set = {str(v) for v in values}
        for pi in pf.PivotItems():
            pi.Visible = str(pi.Value) in values_set
        return json.dumps({"success": True, "pivot": pivot_name, "field": field_name, "visible": values})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def sort_pivot_field(pivot_name: str, field_name: str, order: str = "asc",
                     sheet: str = None, workbook: str = None) -> str:
    """Sort a PivotTable row/column field. order: 'asc' | 'desc'."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        pt = _get_pivot(ws, pivot_name)
        pf = pt.PivotFields(field_name)
        pf.AutoSort(xlAscending if order == "asc" else xlDescending, field_name)
        return json.dumps({"success": True, "pivot": pivot_name, "field": field_name, "order": order})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def add_calculated_field(pivot_name: str, field_name: str, formula: str,
                         sheet: str = None, workbook: str = None) -> str:
    """Add a calculated field to a PivotTable. formula e.g. '=Sales*0.1'."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        pt = _get_pivot(ws, pivot_name)
        pt.CalculatedFields().Add(Name=field_name, Formula=formula)
        return json.dumps({"success": True, "pivot": pivot_name, "field": field_name, "formula": formula})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def list_pivot_fields(pivot_name: str, sheet: str = None, workbook: str = None) -> str:
    """List all fields of a PivotTable with their current orientation."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        pt = _get_pivot(ws, pivot_name)
        orient_map = {xlRowField: "row", xlColumnField: "column",
                      xlPageField: "filter", xlDataField: "value", xlHidden: "hidden"}
        fields = []
        for pf in pt.PivotFields():
            fields.append({"name": pf.Name, "area": orient_map.get(pf.Orientation, "unknown")})
        return json.dumps({"success": True, "pivot": pivot_name, "fields": fields})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def get_pivot_data(pivot_name: str, sheet: str = None, workbook: str = None) -> str:
    """Return the PivotTable data as a 2-D list."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        pt = _get_pivot(ws, pivot_name)
        rng = pt.TableRange2
        raw = rng.Value
        data = []
        for row in (raw if isinstance(raw, tuple) else [[raw]]):
            data.append([str(v) if hasattr(v, "year") else v for v in row])
        return json.dumps({"success": True, "pivot": pivot_name, "data": data,
                           "rows": len(data), "cols": len(data[0]) if data else 0})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def refresh_pivot(pivot_name: str, sheet: str = None, workbook: str = None) -> str:
    """Refresh a PivotTable from its source data."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        pt = _get_pivot(ws, pivot_name)
        pt.RefreshTable()
        return json.dumps({"success": True, "pivot": pivot_name, "refreshed": True})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_pivot_layout(pivot_name: str, layout: str = "tabular", subtotals: bool = True,
                     grand_totals_rows: bool = True, grand_totals_cols: bool = True,
                     sheet: str = None, workbook: str = None) -> str:
    """Set PivotTable layout and summary options.
    layout: 'tabular' | 'outline' | 'compact'
    """
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        pt = _get_pivot(ws, pivot_name)
        layout_map = {"tabular": 2, "outline": 1, "compact": 0}
        for pf in pt.RowFields():
            pf.LayoutForm = layout_map.get(layout, 2)
        pt.RowGrand = grand_totals_rows
        pt.ColumnGrand = grand_totals_cols
        return json.dumps({"success": True, "pivot": pivot_name, "layout": layout})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def delete_pivot(pivot_name: str, sheet: str = None, workbook: str = None) -> str:
    """Delete a PivotTable (retains underlying data range)."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        pt = _get_pivot(ws, pivot_name)
        pt.TableRange2.Clear()
        return json.dumps({"success": True, "deleted": pivot_name})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})
