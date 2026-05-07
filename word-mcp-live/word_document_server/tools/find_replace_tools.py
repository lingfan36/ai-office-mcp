"""Advanced find & replace, word count, and document statistics tools."""

import json
import sys


async def word_live_find_replace_advanced(
    filename: str = None,
    find_text: str = "",
    replace_text: str = "",
    match_case: bool = False,
    whole_word: bool = False,
    use_wildcards: bool = False,
    sounds_like: bool = False,
    all_word_forms: bool = False,
    replace_all: bool = True,
    find_font_name: str = None,
    find_font_size: float = None,
    find_bold: bool = None,
    find_italic: bool = None,
    find_style: str = None,
    replace_font_name: str = None,
    replace_font_size: float = None,
    replace_bold: bool = None,
    replace_italic: bool = None,
    replace_style: str = None,
) -> str:
    """Advanced find & replace in an open Word document via COM.

    Supports wildcards, whole-word matching, case sensitivity, and
    simultaneous formatting changes. Wildcards use Word's own syntax:
      * = any string,  ? = any single character,  [abc] = character class,
      < = word start,  > = word end.

    Args:
        filename: Document name or path (None = active document).
        find_text: Text to find. Use Word wildcard syntax when use_wildcards=True.
        replace_text: Replacement text. Use \\1 \\2 for wildcard capture groups.
        match_case: Case-sensitive search (default False).
        whole_word: Match complete words only (default False).
        use_wildcards: Enable Word wildcard syntax (default False).
        sounds_like: Find words that sound similar (English only).
        all_word_forms: Find all grammatical forms (English only).
        replace_all: Replace all occurrences (default True). False = replace first only.
        find_font_name / find_font_size / find_bold / find_italic: Narrow search by formatting.
        find_style: Find only text with this paragraph style applied.
        replace_font_name / replace_font_size / replace_bold / replace_italic: Apply formatting.
        replace_style: Apply this paragraph style to replaced text.

    Returns:
        JSON with replacement count.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Find/replace tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        with undo_record(app, "MCP: Find & Replace"):
            fd = app.Selection.Find
            fd.ClearFormatting()
            fd.Replacement.ClearFormatting()

            fd.Text = find_text
            fd.MatchCase = match_case
            fd.MatchWholeWord = whole_word
            fd.MatchWildcards = use_wildcards
            fd.MatchSoundsLike = sounds_like
            fd.MatchAllWordForms = all_word_forms
            fd.Forward = True
            fd.Wrap = 1           # wdFindContinue
            fd.Format = bool(
                find_font_name or find_font_size is not None or
                find_bold is not None or find_italic is not None or find_style
            )

            # Find formatting constraints
            if find_font_name:
                fd.Font.Name = find_font_name
            if find_font_size is not None:
                fd.Font.Size = find_font_size
            if find_bold is not None:
                fd.Font.Bold = find_bold
            if find_italic is not None:
                fd.Font.Italic = find_italic
            if find_style:
                fd.Style = doc.Styles(find_style)

            # Replace text
            fd.Replacement.Text = replace_text

            # Replace formatting
            if replace_font_name:
                fd.Replacement.Font.Name = replace_font_name
            if replace_font_size is not None:
                fd.Replacement.Font.Size = replace_font_size
            if replace_bold is not None:
                fd.Replacement.Font.Bold = replace_bold
            if replace_italic is not None:
                fd.Replacement.Font.Italic = replace_italic
            if replace_style:
                fd.Replacement.Style = doc.Styles(replace_style)

            # wdReplaceAll=2, wdReplaceOne=1
            replace_flag = 2 if replace_all else 1
            result = fd.Execute(Replace=replace_flag)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "find_text": find_text,
            "replace_text": replace_text,
            "replaced": result,
            "replace_all": replace_all,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_find_all_occurrences(
    filename: str = None,
    find_text: str = "",
    match_case: bool = False,
    whole_word: bool = False,
    use_wildcards: bool = False,
    max_results: int = 100,
) -> str:
    """Find all occurrences of text in an open Word document and return their positions.

    Args:
        filename: Document name or path (None = active document).
        find_text: Text to search for.
        match_case: Case-sensitive (default False).
        whole_word: Complete words only (default False).
        use_wildcards: Enable Word wildcard syntax (default False).
        max_results: Maximum number of results to return (default 100).

    Returns:
        JSON with list of occurrences including character start/end positions and page numbers.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Find/replace tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        # Use a separate Range to search without moving the selection
        rng = doc.Content
        rng.Find.ClearFormatting()
        rng.Find.Text = find_text
        rng.Find.MatchCase = match_case
        rng.Find.MatchWholeWord = whole_word
        rng.Find.MatchWildcards = use_wildcards
        rng.Find.Forward = True
        rng.Find.Wrap = 0   # wdFindStop

        occurrences = []
        while rng.Find.Execute() and len(occurrences) < max_results:
            page = None
            try:
                page = rng.Information(3)   # wdActiveEndPageNumber=3
            except Exception:
                pass
            # Find which paragraph this is in
            para_idx = None
            try:
                for pi in range(1, doc.Paragraphs.Count + 1):
                    pr = doc.Paragraphs(pi).Range
                    if pr.Start <= rng.Start <= pr.End:
                        para_idx = pi
                        break
            except Exception:
                pass
            occurrences.append({
                "start": rng.Start,
                "end": rng.End,
                "text": rng.Text,
                "page": page,
                "paragraph_index": para_idx,
            })

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "find_text": find_text,
            "count": len(occurrences),
            "occurrences": occurrences,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_get_document_statistics(filename: str = None) -> str:
    """Get word count, character count, paragraph count, and page count for an open document.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with document statistics.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Statistics tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        # wdStatistic constants
        WD_STAT_WORDS = 0
        WD_STAT_LINES = 1
        WD_STAT_PAGES = 2
        WD_STAT_CHARS = 3
        WD_STAT_PARAS = 4
        WD_STAT_CHARS_NO_SPACES = 5

        app = get_word_app()
        doc = find_document(app, filename)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "pages": doc.ComputeStatistics(WD_STAT_PAGES),
            "words": doc.ComputeStatistics(WD_STAT_WORDS),
            "characters": doc.ComputeStatistics(WD_STAT_CHARS),
            "characters_no_spaces": doc.ComputeStatistics(WD_STAT_CHARS_NO_SPACES),
            "paragraphs": doc.ComputeStatistics(WD_STAT_PARAS),
            "lines": doc.ComputeStatistics(WD_STAT_LINES),
            "tables": doc.Tables.Count,
            "images": doc.InlineShapes.Count,
            "shapes": doc.Shapes.Count,
            "sections": doc.Sections.Count,
            "footnotes": doc.Footnotes.Count,
            "endnotes": doc.Endnotes.Count,
            "comments": doc.Comments.Count,
            "fields": doc.Fields.Count,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})
