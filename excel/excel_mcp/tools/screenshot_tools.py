"""Screenshot / visual capture tools (2 operations)."""
import json
import os
import tempfile
from excel_mcp.core.excel_com import get_excel_app, find_workbook, find_sheet


def capture_range(address: str, output_path: str = None, sheet: str = None, workbook: str = None) -> str:
    """Capture a range as a PNG image. Returns base64-encoded PNG if no output_path."""
    try:
        import win32clipboard
        import win32con
        from PIL import Image
        import io, base64

        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        rng = ws.Range(address)
        rng.CopyPicture(Format=2)  # xlBitmap

        win32clipboard.OpenClipboard()
        try:
            data = win32clipboard.GetClipboardData(win32con.CF_DIB)
        finally:
            win32clipboard.CloseClipboard()

        img = Image.open(io.BytesIO(data))
        if output_path:
            img.save(output_path, "PNG")
            return json.dumps({"success": True, "path": output_path, "size": list(img.size)})
        else:
            buf = io.BytesIO()
            img.save(buf, "PNG")
            encoded = base64.b64encode(buf.getvalue()).decode()
            return json.dumps({"success": True, "base64_png": encoded, "size": list(img.size)})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def capture_sheet(output_path: str = None, sheet: str = None, workbook: str = None) -> str:
    """Capture the used range of a worksheet as a PNG image."""
    try:
        app = get_excel_app()
        wb = find_workbook(app, workbook)
        ws = find_sheet(wb, sheet)
        address = ws.UsedRange.Address
        return capture_range(address, output_path=output_path, sheet=ws.Name, workbook=wb.Name)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})
