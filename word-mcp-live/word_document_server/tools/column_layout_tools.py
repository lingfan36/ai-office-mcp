"""Multi-column layout tools for Word sections via COM."""

import json
import sys

_PTS_PER_CM = 28.35


async def word_live_set_column_layout(
    filename: str = None,
    section_index: int = 1,
    columns: int = 2,
    equal_width: bool = True,
    space_between_cm: float = 1.27,
    line_between: bool = False,
    column_widths_cm: list = None,
    space_after_cm: list = None,
) -> str:
    """Set multi-column layout for a section in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        section_index: Section number (1-indexed). Default 1.
        columns: Number of columns (1 = single column, 2 = two columns, etc.).
        equal_width: Make all columns equal width (default True).
            If False, provide column_widths_cm.
        space_between_cm: Space between equal-width columns in cm (default 1.27 cm = 0.5 in).
        line_between: Draw a vertical line between columns (default False).
        column_widths_cm: List of column widths in cm for unequal columns.
            Length must equal columns. Only used when equal_width=False.
        space_after_cm: List of spacer widths after each column in cm.
            Length must equal columns - 1. Only used when equal_width=False.

    Returns:
        JSON with section index, column count, and layout details.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Column layout tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if section_index < 1 or section_index > doc.Sections.Count:
            return json.dumps({
                "error": f"section_index {section_index} out of range (1-{doc.Sections.Count})"
            })

        ps = doc.Sections(section_index).PageSetup

        with undo_record(app, f"MCP: Set Column Layout ({columns} cols)"):
            if equal_width:
                ps.TextColumns.SetCount(columns)
                ps.TextColumns.EvenlySpaced = True
                ps.TextColumns.LineBetween = line_between
                if columns > 1:
                    ps.TextColumns.Spacing = space_between_cm * _PTS_PER_CM
            else:
                if not column_widths_cm or len(column_widths_cm) != columns:
                    return json.dumps({
                        "error": f"column_widths_cm must have exactly {columns} values"
                    })
                spaces = space_after_cm or ([space_between_cm] * (columns - 1))
                ps.TextColumns.SetCount(columns)
                ps.TextColumns.EvenlySpaced = False
                ps.TextColumns.LineBetween = line_between
                for i in range(1, columns + 1):
                    col = ps.TextColumns(i)
                    col.Width = column_widths_cm[i - 1] * _PTS_PER_CM
                    if i < columns:
                        col.SpaceAfter = spaces[i - 1] * _PTS_PER_CM

        result = {
            "success": True,
            "document": doc.Name,
            "section": section_index,
            "columns": columns,
            "equal_width": equal_width,
            "line_between": line_between,
        }
        if equal_width:
            result["space_between_cm"] = space_between_cm
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_get_column_layout(
    filename: str = None,
    section_index: int = 1,
) -> str:
    """Get the current column layout of a section in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        section_index: Section number (1-indexed). Default 1.

    Returns:
        JSON with column count, widths, spacing, and line-between setting.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Column layout tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        if section_index < 1 or section_index > doc.Sections.Count:
            return json.dumps({"error": f"section_index {section_index} out of range"})

        ps = doc.Sections(section_index).PageSetup
        tc = ps.TextColumns
        count = tc.Count

        cols_info = []
        for i in range(1, count + 1):
            c = tc(i)
            cols_info.append({
                "index": i,
                "width_cm": round(c.Width / _PTS_PER_CM, 2),
                "space_after_cm": round(c.SpaceAfter / _PTS_PER_CM, 2) if i < count else 0,
            })

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "section": section_index,
            "columns": count,
            "equal_width": bool(tc.EvenlySpaced),
            "line_between": bool(tc.LineBetween),
            "column_details": cols_info,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_insert_column_break(
    filename: str = None,
    paragraph_index: int = None,
) -> str:
    """Insert a column break at the end of a paragraph in an open Word document.

    The break forces the following text to start in the next column.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index: Insert break after this paragraph (1-indexed).
            None = insert at cursor position (end of active selection).

    Returns:
        JSON confirming insertion.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Column layout tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        with undo_record(app, "MCP: Insert Column Break"):
            if paragraph_index is not None:
                total = doc.Paragraphs.Count
                if paragraph_index < 1 or paragraph_index > total:
                    return json.dumps({"error": f"paragraph_index {paragraph_index} out of range"})
                end = doc.Paragraphs(paragraph_index).Range.End
                rng = doc.Range(end - 1, end - 1)
            else:
                rng = app.Selection.Range
                rng.Collapse(0)  # collapse to end

            rng.InsertBreak(Type=8)  # wdColumnBreak = 8

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "inserted_at_paragraph": paragraph_index,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})
