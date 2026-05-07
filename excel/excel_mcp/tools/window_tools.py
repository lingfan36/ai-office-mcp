"""Window / UI management tools (9 operations)."""
import json
from excel_mcp.core.excel_com import get_excel_app
from excel_mcp.core.constants import xlNormal, xlMinimized, xlMaximized


def show_excel() -> str:
    """Make Excel visible and bring it to the foreground."""
    try:
        app = get_excel_app()
        app.Visible = True
        app.WindowState = xlNormal
        return json.dumps({"success": True, "visible": True})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def hide_excel() -> str:
    """Hide the Excel window."""
    try:
        app = get_excel_app()
        app.Visible = False
        return json.dumps({"success": True, "visible": False})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def get_window_info() -> str:
    """Return current Excel window state, position, and size."""
    try:
        app = get_excel_app()
        state_map = {xlNormal: "normal", xlMinimized: "minimized", xlMaximized: "maximized"}
        return json.dumps({
            "success": True,
            "visible": app.Visible,
            "state": state_map.get(app.WindowState, "unknown"),
            "left": app.Left,
            "top": app.Top,
            "width": app.Width,
            "height": app.Height,
        })
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_window_state(state: str) -> str:
    """Set Excel window state. state: 'normal' | 'minimized' | 'maximized'."""
    try:
        app = get_excel_app()
        state_map = {"normal": xlNormal, "minimized": xlMinimized, "maximized": xlMaximized}
        if state not in state_map:
            raise ValueError(f"Invalid state: {state!r}")
        app.WindowState = state_map[state]
        return json.dumps({"success": True, "state": state})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_window_position(left: float, top: float, width: float = None, height: float = None) -> str:
    """Set Excel window position and optionally size (in points)."""
    try:
        app = get_excel_app()
        app.WindowState = xlNormal
        app.Left = left
        app.Top = top
        if width is not None:
            app.Width = width
        if height is not None:
            app.Height = height
        return json.dumps({"success": True, "left": left, "top": top})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def arrange_window(layout: str) -> str:
    """Apply a preset window layout.
    layout: 'left_half' | 'right_half' | 'top_half' | 'bottom_half' | 'center' | 'full_screen'
    """
    try:
        import win32api
        app = get_excel_app()
        sw = win32api.GetSystemMetrics(0)
        sh = win32api.GetSystemMetrics(1)
        layouts = {
            "left_half":   (0,      0,      sw//2,  sh),
            "right_half":  (sw//2,  0,      sw//2,  sh),
            "top_half":    (0,      0,      sw,     sh//2),
            "bottom_half": (0,      sh//2,  sw,     sh//2),
            "center":      (sw//4,  sh//4,  sw//2,  sh//2),
            "full_screen": (0,      0,      sw,     sh),
        }
        if layout not in layouts:
            raise ValueError(f"Unknown layout: {layout!r}")
        l, t, w, h = layouts[layout]
        app.WindowState = xlNormal
        app.Left, app.Top, app.Width, app.Height = l, t, w, h
        return json.dumps({"success": True, "layout": layout})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def set_status_bar(message: str) -> str:
    """Display a custom message in the Excel status bar."""
    try:
        app = get_excel_app()
        app.StatusBar = message
        return json.dumps({"success": True, "message": message})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


def clear_status_bar() -> str:
    """Restore the default Excel status bar."""
    try:
        app = get_excel_app()
        app.StatusBar = False
        return json.dumps({"success": True, "cleared": True})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})
