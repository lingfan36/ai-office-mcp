"""Document navigation tools — go to page, heading, bookmark, or paragraph via COM."""

import json
import sys


async def word_live_go_to_page(
    filename: str = None,
    page_number: int = 1,
) -> str:
    """Navigate the Word view to a specific page number.

    Args:
        filename: Document name or path (None = active document).
        page_number: 1-based page number to navigate to.

    Returns:
        JSON confirming navigation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Navigation tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)
        doc.Activate()

        # wdGoToPage=1, wdGoToAbsolute=1
        app.Selection.GoTo(What=1, Which=1, Count=page_number)

        # Get actual page after navigation
        actual_page = None
        try:
            actual_page = app.Selection.Information(3)  # wdActiveEndPageNumber
        except Exception:
            pass

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "requested_page": page_number,
            "current_page": actual_page,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_go_to_heading(
    filename: str = None,
    heading_index: int = None,
    heading_text: str = None,
) -> str:
    """Navigate the Word view to a heading paragraph.

    Finds the heading by its order among all headings (heading_index) or by
    partial text match (heading_text). heading_index takes priority.

    Args:
        filename: Document name or path (None = active document).
        heading_index: 1-based index among all heading paragraphs (Heading 1/2/3...).
        heading_text: Partial text to match in a heading paragraph.

    Returns:
        JSON with the heading text and page it is on.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Navigation tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)
        doc.Activate()

        # Collect heading paragraphs
        headings = []
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            style_name = ""
            try:
                style_name = p.Style.NameLocal
            except Exception:
                pass
            if "heading" in style_name.lower() or style_name.lower().startswith("标题"):
                headings.append((i, p))

        if not headings:
            return json.dumps({"error": "No heading paragraphs found in document"})

        target_para = None
        if heading_index is not None:
            if heading_index < 1 or heading_index > len(headings):
                return json.dumps({"error": f"heading_index {heading_index} out of range (1-{len(headings)})"})
            _, target_para = headings[heading_index - 1]
        elif heading_text:
            for _, p in headings:
                try:
                    if heading_text.lower() in p.Range.Text.lower():
                        target_para = p
                        break
                except Exception:
                    pass
            if target_para is None:
                return json.dumps({"error": f"No heading containing '{heading_text}' found"})
        else:
            return json.dumps({"error": "Provide heading_index or heading_text"})

        target_para.Range.Select()
        app.ActiveWindow.ScrollIntoView(target_para.Range)

        heading_txt = ""
        page = None
        try:
            heading_txt = target_para.Range.Text.strip()
            page = target_para.Range.Information(3)
        except Exception:
            pass

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "heading_text": heading_txt,
            "page": page,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_go_to_bookmark(
    filename: str = None,
    bookmark_name: str = "",
) -> str:
    """Navigate the Word view to a named bookmark.

    Args:
        filename: Document name or path (None = active document).
        bookmark_name: Exact name of the bookmark to navigate to.

    Returns:
        JSON with bookmark name and page number.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Navigation tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)
        doc.Activate()

        if not doc.Bookmarks.Exists(bookmark_name):
            available = [doc.Bookmarks(i).Name for i in range(1, doc.Bookmarks.Count + 1)]
            return json.dumps({
                "error": f"Bookmark '{bookmark_name}' not found",
                "available_bookmarks": available,
            })

        bm = doc.Bookmarks(bookmark_name)
        bm.Select()
        app.ActiveWindow.ScrollIntoView(bm.Range)

        page = None
        try:
            page = bm.Range.Information(3)
        except Exception:
            pass

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "bookmark": bookmark_name,
            "page": page,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_scroll_to_paragraph(
    filename: str = None,
    paragraph_index: int = 1,
) -> str:
    """Scroll the Word view so a specific paragraph is visible.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index: 1-based paragraph index.

    Returns:
        JSON with paragraph text preview and page number.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Navigation tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)
        doc.Activate()

        total = doc.Paragraphs.Count
        if paragraph_index < 1 or paragraph_index > total:
            return json.dumps({"error": f"paragraph_index {paragraph_index} out of range (1-{total})"})

        p = doc.Paragraphs(paragraph_index)
        p.Range.Select()
        app.ActiveWindow.ScrollIntoView(p.Range)

        preview = ""
        page = None
        try:
            preview = p.Range.Text.strip()[:80]
            page = p.Range.Information(3)
        except Exception:
            pass

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraph_index": paragraph_index,
            "preview": preview,
            "page": page,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_list_bookmarks(filename: str = None) -> str:
    """List all bookmarks in an open Word document.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with bookmark names, positions, and page numbers.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Navigation tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        bookmarks = []
        for i in range(1, doc.Bookmarks.Count + 1):
            bm = doc.Bookmarks(i)
            page = None
            try:
                page = bm.Range.Information(3)
            except Exception:
                pass
            try:
                bookmarks.append({
                    "index": i,
                    "name": bm.Name,
                    "start": bm.Start,
                    "end": bm.End,
                    "page": page,
                })
            except Exception:
                pass

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "count": len(bookmarks),
            "bookmarks": bookmarks,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})
