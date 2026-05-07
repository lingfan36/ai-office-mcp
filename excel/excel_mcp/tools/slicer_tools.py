"""Slicer tools for PivotTables and Excel Tables (8 operations)."""
import json
from excel_mcp.core.excel_com import get_excel_app, find_workbook, find_sheet, SnapshotUndo


def create_pivot_slicer(pivot_name: str, field_name: str, slicer_name: str = None,
                        left: float = None, top: float = None,
                        width: float = 150, height: float = 200,
                        sheet: str = None, workbook: str = None) -> str:
    """Add a slicer for a PivotTable field."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        pt = None
        for p in ws.PivotTables():
            if p.Name == pivot_name:
                pt = p
                break
        if pt is None:
            raise ValueError(f"PivotTable not found: {pivot_name!r}")
        sc = wb.SlicerCaches.Add2(pt, field_name)
        if slicer_name:
            sc.Name = slicer_name
        slr = sc.Slicers.Add(ws, Left=left or 10, Top=top or 10, Width=width, Height=height)
        return json.dumps({"success": True, "slicer": slr.Name, "cache": sc.Name})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def list_pivot_slicers(pivot_name: str, sheet: str = None, workbook: str = None) -> str:
    """List all slicers connected to a PivotTable."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        results = []
        for sc in wb.SlicerCaches:
            try:
                if sc.PivotTables(1).Name == pivot_name:
                    for slr in sc.Slicers:
                        results.append({"name": slr.Name, "field": sc.SourceName})
            except Exception:
                pass
        return json.dumps({"success": True, "slicers": results})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_pivot_slicer_selection(slicer_cache_name: str, values: list[str], workbook: str = None) -> str:
    """Filter a PivotTable via its slicer cache to show only specified values."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        sc = wb.SlicerCaches(slicer_cache_name)
        values_set = {str(v) for v in values}
        for si in sc.SlicerItems:
            si.Selected = si.Value in values_set
        return json.dumps({"success": True, "cache": slicer_cache_name, "selected": values})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def delete_slicer(slicer_name: str, workbook: str = None) -> str:
    """Delete a slicer and its cache."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        for sc in wb.SlicerCaches:
            for slr in sc.Slicers:
                if slr.Name == slicer_name:
                    slr.Delete()
                    return json.dumps({"success": True, "deleted": slicer_name})
        raise ValueError(f"Slicer not found: {slicer_name!r}")
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def create_table_slicer(table_name: str, column_name: str, slicer_name: str = None,
                        left: float = None, top: float = None,
                        width: float = 150, height: float = 200,
                        sheet: str = None, workbook: str = None) -> str:
    """Add a slicer for an Excel Table column."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        lo = None
        for obj in ws.ListObjects:
            if obj.Name == table_name:
                lo = obj
                break
        if lo is None:
            raise ValueError(f"Table not found: {table_name!r}")
        sc = wb.SlicerCaches.Add2(lo, column_name)
        if slicer_name:
            sc.Name = slicer_name
        slr = sc.Slicers.Add(ws, Left=left or 10, Top=top or 10, Width=width, Height=height)
        return json.dumps({"success": True, "slicer": slr.Name, "cache": sc.Name})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_table_slicer_selection(slicer_cache_name: str, values: list[str], workbook: str = None) -> str:
    """Filter a Table via its slicer cache to show only specified values."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        sc = wb.SlicerCaches(slicer_cache_name)
        values_set = {str(v) for v in values}
        for si in sc.SlicerItems:
            si.Selected = si.Value in values_set
        return json.dumps({"success": True, "cache": slicer_cache_name, "selected": values})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})
