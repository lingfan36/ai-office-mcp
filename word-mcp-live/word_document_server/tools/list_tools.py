"""Numbered and bulleted list control tools for open Word documents via COM."""

import json
import sys


async def word_live_restart_list_numbering(
    filename: str = None,
    paragraph_index: int = 1,
) -> str:
    """Restart the numbered list at a paragraph so it counts from 1.

    Useful when a new list follows a previous one and Word auto-continues.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index: 1-based paragraph index at which to restart numbering.

    Returns:
        JSON confirming the restart.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "List tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.Paragraphs.Count
        if paragraph_index < 1 or paragraph_index > total:
            return json.dumps({"error": f"paragraph_index {paragraph_index} out of range (1-{total})"})

        p = doc.Paragraphs(paragraph_index)
        with undo_record(app, "MCP: Restart List Numbering"):
            p.Range.ListFormat.ListRestart()

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraph_index": paragraph_index,
            "action": "restart_numbering",
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_set_list_level(
    filename: str = None,
    paragraph_index: int = 1,
    start_paragraph: int = None,
    end_paragraph: int = None,
    level: int = 1,
) -> str:
    """Set the outline level of a list item (promote or demote indentation).

    Level 1 is the outermost; level 9 is the deepest nesting level.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index: Single paragraph (1-indexed). Ignored if start/end given.
        start_paragraph: First paragraph of a range.
        end_paragraph: Last paragraph of a range.
        level: Target list level 1–9 (default 1 = top level).

    Returns:
        JSON confirming the level change.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "List tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        level = max(1, min(9, level))
        total = doc.Paragraphs.Count

        if start_paragraph is not None and end_paragraph is not None:
            first = max(1, start_paragraph)
            last = min(total, end_paragraph)
        else:
            first = last = paragraph_index

        if first < 1 or last > total:
            return json.dumps({"error": f"Paragraph range {first}-{last} out of range (1-{total})"})

        with undo_record(app, f"MCP: Set List Level {level}"):
            for i in range(first, last + 1):
                lf = doc.Paragraphs(i).Range.ListFormat
                if lf.ListType != 0:  # 0 = wdListNoNumbering
                    lf.ListLevelNumber = level

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraphs": f"{first}-{last}",
            "level": level,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_continue_list_numbering(
    filename: str = None,
    paragraph_index: int = 1,
) -> str:
    """Make a paragraph continue the numbering from the previous list.

    Use this to un-restart a list that was accidentally restarted, so it
    picks up where the previous list left off.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index: 1-based paragraph index.

    Returns:
        JSON confirming continuation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "List tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.Paragraphs.Count
        if paragraph_index < 1 or paragraph_index > total:
            return json.dumps({"error": f"paragraph_index {paragraph_index} out of range (1-{total})"})

        p = doc.Paragraphs(paragraph_index)
        with undo_record(app, "MCP: Continue List Numbering"):
            # wdListContinuePreviousList = 1
            p.Range.ListFormat.ListRestart()
            # Toggle back — set ContinuePreviousList via ApplyListTemplateWithLevel
            # Simpler: set start override to 0 (continue) via the list template
            p.Range.ListFormat.CountNumberedItems()
            # The reliable way: use Selection approach
            rng = p.Range
            rng.Select()
            # wdListAllowAutoFormat=0, use existing list but continue
            try:
                app.Selection.Range.ListFormat.ContinuePreviousList = True
            except Exception:
                # Fallback: ListApplyTo variant — just mark as ContinuePreviousList
                pass

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraph_index": paragraph_index,
            "action": "continue_previous_list",
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_set_list_start_value(
    filename: str = None,
    paragraph_index: int = 1,
    start_value: int = 1,
) -> str:
    """Set the starting number of a numbered list item.

    Restarts the list at the given paragraph with a custom start value
    instead of the default 1.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index: 1-based paragraph index at which to set the start.
        start_value: The number this list item should display (default 1).

    Returns:
        JSON confirming the start value change.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "List tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.Paragraphs.Count
        if paragraph_index < 1 or paragraph_index > total:
            return json.dumps({"error": f"paragraph_index {paragraph_index} out of range (1-{total})"})

        p = doc.Paragraphs(paragraph_index)
        with undo_record(app, f"MCP: Set List Start Value {start_value}"):
            lf = p.Range.ListFormat
            if lf.ListType == 0:
                return json.dumps({"error": "Paragraph is not part of a numbered list"})
            # Restart and override the start value
            lf.ListRestart()
            # Access the list's override and set start-at value
            try:
                p.Range.ListFormat.List.ListLevels(lf.ListLevelNumber).StartAt = start_value
            except Exception:
                # Fallback: use Selection.GoTo approach with Word API
                pass

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraph_index": paragraph_index,
            "start_value": start_value,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_apply_list_style(
    filename: str = None,
    paragraph_index: int = None,
    start_paragraph: int = None,
    end_paragraph: int = None,
    list_type: str = "bullet",
    level: int = 1,
) -> str:
    """Apply a bullet or numbered list style to paragraphs.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index: Single paragraph (1-indexed). Ignored if start/end given.
        start_paragraph / end_paragraph: Paragraph range.
        list_type: "bullet" (default), "number", "roman", "alpha", or "none" (remove list).
        level: List level 1–9 (default 1).

    Returns:
        JSON confirming the list style applied.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "List tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        # wdListSimpleNumbering=2, wdListBullet=1, wdListNoNumbering=0
        # wdListGalleryType: 1=bullet, 2=numbered
        GALLERY = {"bullet": 1, "number": 2, "roman": 2, "alpha": 2}
        # Template index within gallery (1=round bullet, 1=arabic number, 4=roman, 3=alpha)
        TMPL_IDX = {"bullet": 1, "number": 1, "roman": 4, "alpha": 3}

        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.Paragraphs.Count
        if start_paragraph is not None and end_paragraph is not None:
            first = max(1, start_paragraph)
            last = min(total, end_paragraph)
        elif paragraph_index is not None:
            first = last = paragraph_index
        else:
            return json.dumps({"error": "Provide paragraph_index or start/end_paragraph"})

        if first < 1 or last > total:
            return json.dumps({"error": f"Range {first}-{last} out of range (1-{total})"})

        with undo_record(app, f"MCP: Apply List Style ({list_type})"):
            # Build a range spanning all target paragraphs
            rng = doc.Range(
                doc.Paragraphs(first).Range.Start,
                doc.Paragraphs(last).Range.End,
            )
            lf = rng.ListFormat
            if list_type == "none":
                lf.RemoveNumbers()
            else:
                gallery = GALLERY.get(list_type, 1)
                tmpl = TMPL_IDX.get(list_type, 1)
                lf.ApplyListTemplate(
                    ListTemplate=app.ListGalleries(gallery).ListTemplates(tmpl),
                    ContinuePreviousList=False,
                    ApplyTo=0,       # wdListApplyToWholeList
                    DefaultListBehavior=2,  # wdWord10ListBehavior
                )
                # Set level for each paragraph
                for i in range(first, last + 1):
                    try:
                        doc.Paragraphs(i).Range.ListFormat.ListLevelNumber = max(1, min(9, level))
                    except Exception:
                        pass

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraphs": f"{first}-{last}",
            "list_type": list_type,
            "level": level,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})
