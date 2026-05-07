"""Style management tools — list, apply, create, and modify Word styles via COM."""

import json
import sys


# wdStyleType constants
WD_STYLE_PARAGRAPH = 1
WD_STYLE_CHARACTER = 2
WD_STYLE_TABLE = 3
WD_STYLE_LIST = 4


async def word_live_list_styles(
    filename: str = None,
    style_type: str = "all",
    include_builtin: bool = True,
    include_custom: bool = True,
) -> str:
    """List available styles in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        style_type: Filter by type — "all", "paragraph", "character", "table", "list".
        include_builtin: Include Word's built-in styles (default True).
        include_custom: Include user-defined custom styles (default True).

    Returns:
        JSON array of styles with name, type, builtin flag, and base style.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Style tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        TYPE_FILTER = {
            "all": None,
            "paragraph": WD_STYLE_PARAGRAPH,
            "character": WD_STYLE_CHARACTER,
            "table": WD_STYLE_TABLE,
            "list": WD_STYLE_LIST,
        }
        TYPE_NAME = {
            WD_STYLE_PARAGRAPH: "paragraph",
            WD_STYLE_CHARACTER: "character",
            WD_STYLE_TABLE: "table",
            WD_STYLE_LIST: "list",
        }
        type_filter = TYPE_FILTER.get(style_type)

        app = get_word_app()
        doc = find_document(app, filename)

        styles = []
        for style in doc.Styles:
            try:
                stype = style.Type
                if type_filter is not None and stype != type_filter:
                    continue
                is_builtin = style.BuiltIn
                if is_builtin and not include_builtin:
                    continue
                if not is_builtin and not include_custom:
                    continue

                base_name = None
                try:
                    base_name = style.BaseStyle.NameLocal
                except Exception:
                    pass

                styles.append({
                    "name": style.NameLocal,
                    "type": TYPE_NAME.get(stype, str(stype)),
                    "builtin": is_builtin,
                    "base_style": base_name,
                    "in_use": style.InUse,
                })
            except Exception:
                continue

        styles.sort(key=lambda s: (s["type"], s["name"]))
        return json.dumps({
            "success": True,
            "document": doc.Name,
            "count": len(styles),
            "styles": styles,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_apply_paragraph_style(
    filename: str = None,
    paragraph_index: int = None,
    start_paragraph: int = None,
    end_paragraph: int = None,
    style_name: str = "Normal",
) -> str:
    """Apply a paragraph style to one or more paragraphs in an open Word document.

    Common built-in style names: "Normal", "Heading 1" – "Heading 9",
    "Title", "Subtitle", "Quote", "Intense Quote", "List Paragraph",
    "Body Text", "Caption", "No Spacing".

    Args:
        filename: Document name or path (None = active document).
        paragraph_index: Single paragraph (1-indexed).
        start_paragraph / end_paragraph: Range of paragraphs.
        style_name: Style name as it appears in Word (localised). Default "Normal".

    Returns:
        JSON with count of paragraphs updated.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Style tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)
        total = doc.Paragraphs.Count

        if start_paragraph is not None and end_paragraph is not None:
            indices = range(max(1, start_paragraph), min(end_paragraph + 1, total + 1))
        elif paragraph_index is not None:
            if paragraph_index < 1 or paragraph_index > total:
                return json.dumps({"error": f"paragraph_index {paragraph_index} out of range"})
            indices = [paragraph_index]
        else:
            indices = range(1, total + 1)

        with undo_record(app, f"MCP: Apply Style {style_name}"):
            for i in indices:
                doc.Paragraphs(i).Style = doc.Styles(style_name)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "style_applied": style_name,
            "paragraphs_affected": len(list(indices)),
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_apply_character_style(
    filename: str = None,
    start: int = None,
    end: int = None,
    start_paragraph: int = None,
    end_paragraph: int = None,
    style_name: str = "Default Paragraph Font",
) -> str:
    """Apply a character style to a text range in an open Word document.

    Common character styles: "Strong", "Emphasis", "Intense Emphasis",
    "Book Title", "Default Paragraph Font", "Hyperlink".

    Args:
        filename: Document name or path (None = active document).
        start / end: Character positions.
        start_paragraph / end_paragraph: Alternative — operate on full paragraphs.
        style_name: Character style name (default "Default Paragraph Font").

    Returns:
        JSON confirming style applied.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Style tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if start_paragraph is not None:
            ep = end_paragraph if end_paragraph is not None else start_paragraph
            rng = doc.Range(
                doc.Paragraphs(start_paragraph).Range.Start,
                doc.Paragraphs(ep).Range.End,
            )
        elif start is not None and end is not None:
            rng = doc.Range(start, end)
        else:
            return json.dumps({"error": "Provide start/end or start_paragraph/end_paragraph"})

        with undo_record(app, f"MCP: Apply Char Style {style_name}"):
            rng.Style = doc.Styles(style_name)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "style_applied": style_name,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_create_style(
    filename: str = None,
    style_name: str = "",
    style_type: str = "paragraph",
    base_style: str = "Normal",
    font_name: str = None,
    font_size: float = None,
    bold: bool = None,
    italic: bool = None,
    font_color: str = None,
    alignment: str = None,
    space_before_pt: float = None,
    space_after_pt: float = None,
    left_indent_cm: float = None,
) -> str:
    """Create a new custom style in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        style_name: Name for the new style (must be unique).
        style_type: "paragraph" or "character".
        base_style: Existing style to inherit from (default "Normal").
        font_name: Font family for this style.
        font_size: Font size in points.
        bold / italic: Bold/italic for this style.
        font_color: Font color as "#RRGGBB".
        alignment: "left", "center", "right", "justify" (paragraph styles only).
        space_before_pt / space_after_pt: Paragraph spacing (paragraph styles only).
        left_indent_cm: Left indent in centimetres (paragraph styles only).

    Returns:
        JSON confirming style creation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Style tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        TYPE_MAP = {"paragraph": WD_STYLE_PARAGRAPH, "character": WD_STYLE_CHARACTER}
        ALIGN_MAP = {"left": 0, "center": 1, "right": 2, "justify": 3}

        app = get_word_app()
        doc = find_document(app, filename)

        if not style_name:
            return json.dumps({"error": "style_name is required"})

        style = doc.Styles.Add(style_name, TYPE_MAP.get(style_type, WD_STYLE_PARAGRAPH))
        try:
            style.BaseStyle = doc.Styles(base_style)
        except Exception:
            pass

        font = style.Font
        if font_name:
            font.Name = font_name
        if font_size:
            font.Size = font_size
        if bold is not None:
            font.Bold = bold
        if italic is not None:
            font.Italic = italic
        if font_color:
            c = font_color.lstrip("#")
            font.Color = int(c[0:2], 16) + (int(c[2:4], 16) << 8) + (int(c[4:6], 16) << 16)

        if style_type == "paragraph":
            pf = style.ParagraphFormat
            if alignment and alignment in ALIGN_MAP:
                pf.Alignment = ALIGN_MAP[alignment]
            if space_before_pt is not None:
                pf.SpaceBefore = space_before_pt
            if space_after_pt is not None:
                pf.SpaceAfter = space_after_pt
            if left_indent_cm is not None:
                pf.LeftIndent = left_indent_cm * 28.35

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "style_name": style.NameLocal,
            "style_type": style_type,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_delete_style(
    filename: str = None,
    style_name: str = "",
) -> str:
    """Delete a custom style from an open Word document.

    Note: Built-in Word styles cannot be deleted.

    Args:
        filename: Document name or path (None = active document).
        style_name: Name of the custom style to delete.

    Returns:
        JSON confirming deletion.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Style tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        style = doc.Styles(style_name)
        if style.BuiltIn:
            return json.dumps({"error": f"Cannot delete built-in style '{style_name}'"})
        style.Delete()
        return json.dumps({"success": True, "deleted": style_name})
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_rename_style(
    filename: str = None,
    old_name: str = "",
    new_name: str = "",
) -> str:
    """Rename a custom style in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        old_name: Current style name.
        new_name: New style name.

    Returns:
        JSON confirming rename.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Style tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        style = doc.Styles(old_name)
        if style.BuiltIn:
            return json.dumps({"error": f"Cannot rename built-in style '{old_name}'"})
        style.NameLocal = new_name
        return json.dumps({"success": True, "old_name": old_name, "new_name": new_name})
    except Exception as e:
        return json.dumps({"error": str(e)})
