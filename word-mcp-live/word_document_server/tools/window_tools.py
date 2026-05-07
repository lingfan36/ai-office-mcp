"""Word application window management tools."""

import json
import sys

# Word window state constants
WD_WINDOW_STATE_NORMAL = 0
WD_WINDOW_STATE_MAXIMIZE = 1
WD_WINDOW_STATE_MINIMIZE = 2


async def word_show(filename: str = None) -> str:
    """Make the Word application window visible and bring it to the foreground.

    Args:
        filename: If provided, also activate this specific document.

    Returns:
        JSON confirming visibility.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Window tools are only available on Windows"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app(visible=True)
        app.Visible = True
        try:
            app.WindowState = WD_WINDOW_STATE_NORMAL
        except Exception:
            pass
        app.Activate()
        if filename:
            doc = find_document(app, filename)
            doc.Activate()
        return json.dumps({"success": True, "visible": True})
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_hide() -> str:
    """Hide the Word application window (Word keeps running in background).

    Returns:
        JSON confirming hidden state.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Window tools are only available on Windows"})
    try:
        from word_document_server.core.word_com import get_word_app

        app = get_word_app(auto_launch=False)
        app.Visible = False
        return json.dumps({"success": True, "visible": False})
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_get_window_info() -> str:
    """Get current Word window state, position, and size.

    Returns:
        JSON with visible, state (normal/minimized/maximized), left, top, width, height.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Window tools are only available on Windows"})
    try:
        from word_document_server.core.word_com import get_word_app

        app = get_word_app(auto_launch=False)

        state_map = {
            WD_WINDOW_STATE_NORMAL: "normal",
            WD_WINDOW_STATE_MAXIMIZE: "maximized",
            WD_WINDOW_STATE_MINIMIZE: "minimized",
        }
        state_val = WD_WINDOW_STATE_NORMAL
        try:
            state_val = app.WindowState
        except Exception:
            pass

        return json.dumps({
            "success": True,
            "visible": app.Visible,
            "state": state_map.get(state_val, "unknown"),
            "left": app.Left,
            "top": app.Top,
            "width": app.Width,
            "height": app.Height,
            "active_document": app.ActiveDocument.Name if app.Documents.Count > 0 else None,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_set_window_state(state: str) -> str:
    """Set the Word window state.

    Args:
        state: "normal", "minimized", or "maximized".

    Returns:
        JSON confirming the new state.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Window tools are only available on Windows"})
    try:
        from word_document_server.core.word_com import get_word_app

        STATE_MAP = {
            "normal": WD_WINDOW_STATE_NORMAL,
            "minimized": WD_WINDOW_STATE_MINIMIZE,
            "maximized": WD_WINDOW_STATE_MAXIMIZE,
        }
        if state not in STATE_MAP:
            return json.dumps({"error": f"Invalid state '{state}'. Use: normal, minimized, maximized"})

        app = get_word_app(auto_launch=False)
        app.Visible = True
        app.WindowState = STATE_MAP[state]
        return json.dumps({"success": True, "state": state})
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_set_window_position(
    left: float,
    top: float,
    width: float = None,
    height: float = None,
) -> str:
    """Set Word window position and optionally size (in points).

    Args:
        left: Distance from left edge of screen in points.
        top: Distance from top edge of screen in points.
        width: Window width in points (optional).
        height: Window height in points (optional).

    Returns:
        JSON confirming new position.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Window tools are only available on Windows"})
    try:
        from word_document_server.core.word_com import get_word_app

        app = get_word_app(auto_launch=False)
        app.Visible = True
        app.WindowState = WD_WINDOW_STATE_NORMAL
        app.Left = left
        app.Top = top
        if width is not None:
            app.Width = width
        if height is not None:
            app.Height = height
        return json.dumps({
            "success": True,
            "left": app.Left,
            "top": app.Top,
            "width": app.Width,
            "height": app.Height,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_set_status_bar(text: str) -> str:
    """Display a message in the Word status bar.

    Args:
        text: Message to show in the status bar.

    Returns:
        JSON confirming the message was set.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Window tools are only available on Windows"})
    try:
        from word_document_server.core.word_com import get_word_app

        app = get_word_app(auto_launch=False)
        app.StatusBar = text
        return json.dumps({"success": True, "text": text})
    except Exception as e:
        return json.dumps({"error": str(e)})
