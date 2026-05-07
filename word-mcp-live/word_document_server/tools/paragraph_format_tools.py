"""Paragraph-level formatting tools — indent, tabs, borders, shading, character spacing."""

import json
import sys
from typing import Optional

_PTS_PER_CM = 28.35
_PTS_PER_INCH = 72.0


def _resolve_para_range(doc, paragraph_index, start_paragraph, end_paragraph):
    """Return list of 1-based paragraph indices to operate on."""
    total = doc.Paragraphs.Count
    if start_paragraph is not None and end_paragraph is not None:
        return list(range(max(1, start_paragraph), min(end_paragraph + 1, total + 1)))
    if paragraph_index is not None:
        if paragraph_index < 1 or paragraph_index > total:
            raise ValueError(f"paragraph_index {paragraph_index} out of range (1-{total})")
        return [paragraph_index]
    return list(range(1, total + 1))


async def word_live_set_paragraph_indent(
    filename: str = None,
    paragraph_index: int = None,
    start_paragraph: int = None,
    end_paragraph: int = None,
    left_indent_cm: float = None,
    right_indent_cm: float = None,
    first_line_indent_cm: float = None,
    hanging_indent_cm: float = None,
) -> str:
    """Set paragraph indentation in an open Word document.

    Provide exactly one of paragraph_index or start_paragraph+end_paragraph.
    Omit all three to apply to the entire document.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index: Single paragraph (1-indexed).
        start_paragraph: Start of range (1-indexed, inclusive).
        end_paragraph: End of range (1-indexed, inclusive).
        left_indent_cm: Left indent in centimetres (0 = no indent).
        right_indent_cm: Right indent in centimetres.
        first_line_indent_cm: First-line indent in centimetres (positive = indent in).
        hanging_indent_cm: Hanging indent in centimetres (positive = hanging out).
            Sets first_line to negative of this value — do not combine with first_line_indent_cm.

    Returns:
        JSON with count of affected paragraphs.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Paragraph format tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)
        indices = _resolve_para_range(doc, paragraph_index, start_paragraph, end_paragraph)

        with undo_record(app, "MCP: Set Paragraph Indent"):
            for i in indices:
                pf = doc.Paragraphs(i).Format
                if left_indent_cm is not None:
                    pf.LeftIndent = left_indent_cm * _PTS_PER_CM
                if right_indent_cm is not None:
                    pf.RightIndent = right_indent_cm * _PTS_PER_CM
                if hanging_indent_cm is not None:
                    pf.LeftIndent = (pf.LeftIndent or 0) + hanging_indent_cm * _PTS_PER_CM
                    pf.FirstLineIndent = -hanging_indent_cm * _PTS_PER_CM
                elif first_line_indent_cm is not None:
                    pf.FirstLineIndent = first_line_indent_cm * _PTS_PER_CM

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraphs_affected": len(indices),
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_set_tab_stops(
    filename: str = None,
    paragraph_index: int = None,
    start_paragraph: int = None,
    end_paragraph: int = None,
    tabs: list = None,
    clear_existing: bool = False,
) -> str:
    """Set custom tab stops on paragraphs in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index: Single paragraph (1-indexed).
        start_paragraph / end_paragraph: Range of paragraphs.
        tabs: List of tab stop dicts, each with:
            - position_cm: Tab position in centimetres (required).
            - alignment: "left" | "center" | "right" | "decimal" | "bar" (default "left").
            - leader: "none" | "dots" | "dashes" | "line" (default "none").
        clear_existing: If True, remove all existing tab stops first.

    Returns:
        JSON with count of affected paragraphs.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Paragraph format tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        # wdTab* constants
        ALIGN_MAP = {"left": 0, "center": 1, "right": 2, "decimal": 3, "bar": 4}
        LEADER_MAP = {"none": 0, "dots": 1, "dashes": 2, "line": 3}

        app = get_word_app()
        doc = find_document(app, filename)
        indices = _resolve_para_range(doc, paragraph_index, start_paragraph, end_paragraph)

        parsed_tabs = []
        for t in (tabs or []):
            pos = t.get("position_cm", 0) * _PTS_PER_CM
            align = ALIGN_MAP.get(t.get("alignment", "left"), 0)
            leader = LEADER_MAP.get(t.get("leader", "none"), 0)
            parsed_tabs.append((pos, align, leader))

        with undo_record(app, "MCP: Set Tab Stops"):
            for i in indices:
                pf = doc.Paragraphs(i).Format
                if clear_existing:
                    pf.TabStops.ClearAll()
                for pos, align, leader in parsed_tabs:
                    pf.TabStops.Add(Position=pos, Alignment=align, Leader=leader)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraphs_affected": len(indices),
            "tabs_set": len(parsed_tabs),
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_set_paragraph_border(
    filename: str = None,
    paragraph_index: int = None,
    start_paragraph: int = None,
    end_paragraph: int = None,
    sides: list = None,
    line_style: str = "single",
    line_width_pt: float = 0.75,
    color: str = "#000000",
) -> str:
    """Add borders to paragraphs in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index / start_paragraph / end_paragraph: Target paragraphs.
        sides: List of sides to border — any of "top", "bottom", "left", "right", "all".
               Default ["all"].
        line_style: Border style — "single", "double", "dotted", "dashed", "thick" (default "single").
        line_width_pt: Border width in points (default 0.75).
        color: Border color as "#RRGGBB" hex (default black).

    Returns:
        JSON with count of affected paragraphs.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Paragraph format tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        # wdBorderType: Top=1, Bottom=-1, Left=2, Right=-2
        SIDE_MAP = {"top": 1, "bottom": -1, "left": 2, "right": -2}
        STYLE_MAP = {"single": 1, "thick": 2, "double": 3, "dotted": 4, "dashed": 5}
        # wdLineWidth in eighths of a point: 0.75pt = 6, 1pt = 8, 1.5pt = 12, 2.25pt = 18
        width_map = {0.75: 6, 1.0: 8, 1.5: 12, 2.25: 18}
        width_const = min(width_map, key=lambda k: abs(k - line_width_pt))
        width_val = width_map[width_const]

        c = color.lstrip("#")
        rgb_val = int(c[0:2], 16) + (int(c[2:4], 16) << 8) + (int(c[4:6], 16) << 16)

        if not sides or "all" in sides:
            target_sides = list(SIDE_MAP.keys())
        else:
            target_sides = [s for s in sides if s in SIDE_MAP]

        app = get_word_app()
        doc = find_document(app, filename)
        indices = _resolve_para_range(doc, paragraph_index, start_paragraph, end_paragraph)

        with undo_record(app, "MCP: Set Paragraph Border"):
            for i in indices:
                borders = doc.Paragraphs(i).Borders
                for side in target_sides:
                    b = borders(SIDE_MAP[side])
                    b.LineStyle = STYLE_MAP.get(line_style, 1)
                    b.LineWidth = width_val
                    b.Color = rgb_val

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraphs_affected": len(indices),
            "sides": target_sides,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_clear_paragraph_border(
    filename: str = None,
    paragraph_index: int = None,
    start_paragraph: int = None,
    end_paragraph: int = None,
) -> str:
    """Remove all borders from paragraphs in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index / start_paragraph / end_paragraph: Target paragraphs.

    Returns:
        JSON with count of affected paragraphs.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Paragraph format tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)
        indices = _resolve_para_range(doc, paragraph_index, start_paragraph, end_paragraph)

        with undo_record(app, "MCP: Clear Paragraph Border"):
            for i in indices:
                doc.Paragraphs(i).Borders.ClearFormatting()

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraphs_affected": len(indices),
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_set_paragraph_shading(
    filename: str = None,
    paragraph_index: int = None,
    start_paragraph: int = None,
    end_paragraph: int = None,
    background_color: str = None,
    foreground_color: str = None,
    pattern: str = "solid",
) -> str:
    """Set paragraph background shading (fill color) in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index / start_paragraph / end_paragraph: Target paragraphs.
        background_color: Background fill color as "#RRGGBB" (e.g. "#FFFF00" for yellow).
                          Use "#FFFFFF" for white, or None to keep current.
        foreground_color: Foreground pattern color as "#RRGGBB" (default black).
        pattern: Shading pattern — "solid" (default), "clear" (remove shading), "5_percent",
                 "10_percent", "20_percent", "25_percent", "50_percent".

    Returns:
        JSON with count of affected paragraphs.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Paragraph format tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        # wdTextureIndex: wdTextureNone=0, wdTextureSolid=-1, 5%=250, 10%=251, etc.
        PATTERN_MAP = {
            "clear": 0,          # wdTextureNone
            "solid": -1,         # wdTextureSolid
            "5_percent": 250,
            "10_percent": 251,
            "20_percent": 253,
            "25_percent": 254,
            "50_percent": 256,
        }

        def hex_to_bgr(h):
            h = h.lstrip("#")
            return int(h[0:2], 16) + (int(h[2:4], 16) << 8) + (int(h[4:6], 16) << 16)

        app = get_word_app()
        doc = find_document(app, filename)
        indices = _resolve_para_range(doc, paragraph_index, start_paragraph, end_paragraph)

        with undo_record(app, "MCP: Set Paragraph Shading"):
            for i in indices:
                shading = doc.Paragraphs(i).Shading
                shading.Texture = PATTERN_MAP.get(pattern, -1)
                if background_color:
                    shading.BackgroundPatternColor = hex_to_bgr(background_color)
                if foreground_color:
                    shading.ForegroundPatternColor = hex_to_bgr(foreground_color)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraphs_affected": len(indices),
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_set_character_spacing(
    filename: str = None,
    start: int = None,
    end: int = None,
    start_paragraph: int = None,
    end_paragraph: int = None,
    spacing_pt: float = None,
    scale_percent: int = None,
    position_pt: float = None,
    kerning_pt: float = None,
) -> str:
    """Set character spacing properties on a text range in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        start / end: Character positions (use word_live_find_text to locate).
        start_paragraph / end_paragraph: Alternative — operate on full paragraphs.
        spacing_pt: Letter spacing in points. Positive = expanded, negative = condensed.
                    0 = normal.
        scale_percent: Horizontal scaling as percentage (e.g. 90, 100, 150).
        position_pt: Vertical position offset in points.
                     Positive = raised (superscript-like), negative = lowered.
        kerning_pt: Enable/disable kerning for font sizes above this point size.
                    0 = disable kerning.

    Returns:
        JSON confirming formatting applied.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Paragraph format tools are Windows only"})
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

        with undo_record(app, "MCP: Set Character Spacing"):
            font = rng.Font
            if spacing_pt is not None:
                font.Spacing = spacing_pt
            if scale_percent is not None:
                font.Scaling = scale_percent
            if position_pt is not None:
                font.Position = position_pt
            if kerning_pt is not None:
                if kerning_pt <= 0:
                    font.Kerning = 0
                else:
                    font.Kerning = kerning_pt

        return json.dumps({"success": True, "document": doc.Name})
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_set_text_effects(
    filename: str = None,
    start: int = None,
    end: int = None,
    start_paragraph: int = None,
    end_paragraph: int = None,
    superscript: bool = None,
    subscript: bool = None,
    small_caps: bool = None,
    all_caps: bool = None,
    hidden: bool = None,
    double_strikethrough: bool = None,
    outline: bool = None,
    shadow: bool = None,
    emboss: bool = None,
    engrave: bool = None,
) -> str:
    """Apply advanced text effects to a range in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        start / end: Character positions.
        start_paragraph / end_paragraph: Alternative — operate on full paragraphs.
        superscript: Superscript on/off.
        subscript: Subscript on/off.
        small_caps: Small capitals on/off.
        all_caps: All capitals on/off.
        hidden: Hidden text on/off.
        double_strikethrough: Double strikethrough on/off.
        outline: Outline effect on/off.
        shadow: Shadow effect on/off.
        emboss: Emboss effect on/off.
        engrave: Engrave (imprint) effect on/off.

    Returns:
        JSON confirming effects applied.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Paragraph format tools are Windows only"})
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

        effects = []
        with undo_record(app, "MCP: Set Text Effects"):
            font = rng.Font
            if superscript is not None:
                font.Superscript = superscript
                effects.append(f"superscript={superscript}")
            if subscript is not None:
                font.Subscript = subscript
                effects.append(f"subscript={subscript}")
            if small_caps is not None:
                font.SmallCaps = small_caps
                effects.append(f"small_caps={small_caps}")
            if all_caps is not None:
                font.AllCaps = all_caps
                effects.append(f"all_caps={all_caps}")
            if hidden is not None:
                font.Hidden = hidden
                effects.append(f"hidden={hidden}")
            if double_strikethrough is not None:
                font.DoubleStrikeThrough = double_strikethrough
                effects.append(f"double_strikethrough={double_strikethrough}")
            if outline is not None:
                font.Outline = outline
                effects.append(f"outline={outline}")
            if shadow is not None:
                font.Shadow = shadow
                effects.append(f"shadow={shadow}")
            if emboss is not None:
                font.Emboss = emboss
                effects.append(f"emboss={emboss}")
            if engrave is not None:
                font.Engrave = engrave
                effects.append(f"engrave={engrave}")

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "effects_applied": effects,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_clear_formatting(
    filename: str = None,
    start: int = None,
    end: int = None,
    start_paragraph: int = None,
    end_paragraph: int = None,
    clear_char_formatting: bool = True,
    clear_para_formatting: bool = True,
) -> str:
    """Clear all direct formatting from a range, resetting to the applied style defaults.

    Args:
        filename: Document name or path (None = active document).
        start / end: Character positions.
        start_paragraph / end_paragraph: Alternative — operate on full paragraphs.
        clear_char_formatting: Reset font/color/bold/italic etc. (default True).
        clear_para_formatting: Reset alignment/indent/spacing etc. (default True).

    Returns:
        JSON confirming formatting cleared.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Paragraph format tools are Windows only"})
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

        with undo_record(app, "MCP: Clear Formatting"):
            if clear_char_formatting:
                rng.Font.Reset()
            if clear_para_formatting:
                for p in rng.Paragraphs:
                    p.Format.Reset()

        return json.dumps({"success": True, "document": doc.Name})
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_get_paragraph_formatting(
    filename: str = None,
    paragraph_index: int = 1,
) -> str:
    """Inspect all formatting properties of a paragraph in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index: Paragraph to inspect (1-indexed, default 1).

    Returns:
        JSON with font, alignment, indent, spacing, borders, and style info.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Paragraph format tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.Paragraphs.Count
        if paragraph_index < 1 or paragraph_index > total:
            return json.dumps({
                "error": f"paragraph_index {paragraph_index} out of range (1-{total})"
            })

        para = doc.Paragraphs(paragraph_index)
        pf = para.Format
        rng = para.Range
        font = rng.Font

        ALIGN_MAP = {0: "left", 1: "center", 2: "right", 3: "justify", 4: "distributed"}
        LINE_RULE = {0: "single", 1: "1.5_lines", 2: "double", 3: "at_least", 4: "exactly", 5: "multiple"}

        def safe(v):
            try:
                return v if v != 9999999 else None
            except Exception:
                return None

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraph_index": paragraph_index,
            "text_preview": rng.Text[:80] if rng.Text else "",
            "style": para.Style.NameLocal,
            "alignment": ALIGN_MAP.get(safe(pf.Alignment), str(safe(pf.Alignment))),
            "indent": {
                "left_pt": safe(pf.LeftIndent),
                "right_pt": safe(pf.RightIndent),
                "first_line_pt": safe(pf.FirstLineIndent),
            },
            "spacing": {
                "before_pt": safe(pf.SpaceBefore),
                "after_pt": safe(pf.SpaceAfter),
                "line_spacing_pt": safe(pf.LineSpacing),
                "line_spacing_rule": LINE_RULE.get(safe(pf.LineSpacingRule), str(safe(pf.LineSpacingRule))),
            },
            "pagination": {
                "keep_with_next": safe(pf.KeepWithNext),
                "keep_together": safe(pf.KeepTogether),
                "page_break_before": safe(pf.PageBreakBefore),
                "widow_control": safe(pf.WidowControl),
            },
            "font": {
                "name": safe(font.Name),
                "size_pt": safe(font.Size),
                "bold": safe(font.Bold),
                "italic": safe(font.Italic),
                "underline": safe(font.Underline),
                "strikethrough": safe(font.StrikeThrough),
                "color_bgr": safe(font.Color),
            },
        })
    except Exception as e:
        return json.dumps({"error": str(e)})
