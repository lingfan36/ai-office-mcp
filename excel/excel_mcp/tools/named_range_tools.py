"""Named Range tools (6 operations)."""
import json
from excel_mcp.core.excel_com import get_excel_app, find_workbook, find_sheet, SnapshotUndo


def list_named_ranges(workbook: str = None) -> str:
    """List all named ranges in a workbook."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        names = [{"name": n.Name, "refers_to": n.RefersTo,
                  "comment": n.Comment or ""} for n in wb.Names]
        return json.dumps({"success": True, "named_ranges": names, "count": len(names)})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def read_named_range(name: str, workbook: str = None) -> str:
    """Read the value(s) of a named range."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        rng = wb.Names(name).RefersToRange
        raw = rng.Value
        if isinstance(raw, tuple):
            data = [list(row) for row in raw]
        else:
            data = [[raw]]
        return json.dumps({"success": True, "name": name, "value": data,
                           "address": rng.Address})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def write_named_range(name: str, value, workbook: str = None) -> str:
    """Write a value (or 2-D list) into a named range."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        wb.Names(name).RefersToRange.Value = value
        return json.dumps({"success": True, "name": name})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def create_named_range(name: str, address: str, sheet: str = None,
                       comment: str = None, workbook: str = None) -> str:
    """Create a new named range. address should include sheet if cross-sheet."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        ref = f"={ws.Name}!{address}" if sheet else f"={address}"
        n = wb.Names.Add(Name=name, RefersTo=ref)
        if comment:
            n.Comment = comment
        return json.dumps({"success": True, "name": name, "refers_to": n.RefersTo})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def update_named_range(name: str, new_address: str, sheet: str = None, workbook: str = None) -> str:
    """Change the range an existing name refers to."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        ws = find_sheet(wb, sheet)
        ref = f"={ws.Name}!{new_address}" if sheet else f"={new_address}"
        wb.Names(name).RefersTo = ref
        return json.dumps({"success": True, "name": name, "new_refers_to": ref})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def delete_named_range(name: str, workbook: str = None) -> str:
    """Delete a named range."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        SnapshotUndo.push(wb)
        wb.Names(name).Delete()
        return json.dumps({"success": True, "deleted": name})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})
