"""Image insertion and shape management."""
import json
import os
from excel_mcp.core.excel_com import get_excel_app, find_workbook, find_sheet


def insert_image(image_path: str, anchor_cell: str, width: float = None,
                 height: float = None, sheet: str = None, workbook: str = None) -> str:
    """Insert an image anchored to the top-left of *anchor_cell*.

    *width* and *height* are in points. If only one is given the other is
    scaled proportionally. If neither is given the image keeps its original size.
    """
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        abs_path = os.path.abspath(image_path)
        cell = ws.Range(anchor_cell)
        left = cell.Left
        top = cell.Top
        # Use -1 to let Excel/WPS determine size from the file initially.
        shape = ws.Shapes.AddPicture(
            Filename=abs_path,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=left,
            Top=top,
            Width=-1,
            Height=-1,
        )
        orig_w = shape.Width
        orig_h = shape.Height
        if width and height:
            shape.Width = width
            shape.Height = height
        elif width:
            shape.Width = width
            shape.Height = orig_h * (width / orig_w)
        elif height:
            shape.Height = height
            shape.Width = orig_w * (height / orig_h)
        return json.dumps({
            "success": True,
            "name": shape.Name,
            "left": shape.Left,
            "top": shape.Top,
            "width": shape.Width,
            "height": shape.Height,
        })
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def list_images(sheet: str = None, workbook: str = None) -> str:
    """List all picture shapes on a sheet."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        pics = []
        for s in ws.Shapes:
            # Type 13 = msoPicture; also catch linked pictures (type 11)
            if s.Type in (11, 13):
                pics.append({
                    "name": s.Name,
                    "left": s.Left,
                    "top": s.Top,
                    "width": s.Width,
                    "height": s.Height,
                })
        return json.dumps({"success": True, "images": pics, "count": len(pics)})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def delete_image(name: str, sheet: str = None, workbook: str = None) -> str:
    """Delete an image/shape by its name."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        ws.Shapes(name).Delete()
        return json.dumps({"success": True, "deleted": name})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})
