"""Cell comment (note) operations."""
import json
from excel_mcp.core.excel_com import get_excel_app, find_workbook, find_sheet


def add_comment(address: str, text: str, author: str = None,
                sheet: str = None, workbook: str = None) -> str:
    """Add (or replace) a comment on a cell."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        cell = ws.Range(address)
        try:
            cell.Comment.Delete()
        except Exception:
            pass
        comment = cell.AddComment(text)
        if author:
            try:
                comment.Author = author
            except Exception:
                pass
        return json.dumps({"success": True, "address": address, "text": text})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def get_comments(sheet: str = None, workbook: str = None) -> str:
    """Return all comments on a sheet."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        result = []
        for c in ws.Comments:
            result.append({
                "address": c.Parent.Address,
                "text": c.Text(),
                "author": c.Author,
                "visible": c.Visible,
            })
        return json.dumps({"success": True, "comments": result, "count": len(result)})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def delete_comment(address: str, sheet: str = None, workbook: str = None) -> str:
    """Delete the comment on a cell (no-op if no comment exists)."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        try:
            ws.Range(address).Comment.Delete()
        except Exception:
            pass
        return json.dumps({"success": True, "address": address})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def show_hide_comment(address: str, visible: bool = True,
                      sheet: str = None, workbook: str = None) -> str:
    """Show or hide a cell comment."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        ws.Range(address).Comment.Visible = visible
        return json.dumps({"success": True, "address": address, "visible": visible})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})
