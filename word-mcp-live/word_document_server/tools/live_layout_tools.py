"""COM-based layout tools for Microsoft Word.

These tools operate on documents currently open in Word via COM automation.
They provide layout, header/footer, spacing, bookmark, watermark, and section
management for files that are open (and locked) in Word.
"""

import json
import sys

# macOS JXA dispatch
_MAC_AVAILABLE = sys.platform == 'darwin'

# 1 inch = 72 points (avoid app.InchesToPoints which can fail on some COM setups)
_PTS_PER_INCH = 72.0
_PTS_PER_CM = 72.0 / 2.54  # ≈ 28.346 pt/cm

# Sanity ceiling for any single dimension in points (≈ 56 inches).
# Word's hard limits are 22 in for page, 31.5 in for margin/gutter — this
# generous ceiling catches obvious unit-mistake bugs (e.g. someone passing
# 2540 thinking they're in twentieths of a point).
_MAX_DIMENSION_PT = 4032.0


def _resolve_pt(cm: float = None, inches: float = None, label: str = "") -> tuple[float | None, str]:
    """Pick whichever of cm / inches is given, convert to points, validate.

    Returns (points, source_unit_for_logging). If both given, cm wins and
    we attach a note to the log. If neither given, returns (None, '').
    """
    if cm is None and inches is None:
        return None, ""
    if cm is not None and inches is not None:
        # Prefer cm (regression-safe: callers who set *_cm explicitly want cm)
        pt = float(cm) * _PTS_PER_CM
        return pt, f"{label}={cm}cm (inches param ignored — both supplied)"
    if cm is not None:
        if cm < 0 or cm * _PTS_PER_CM > _MAX_DIMENSION_PT:
            raise ValueError(f"{label}={cm}cm is out of plausible range")
        return float(cm) * _PTS_PER_CM, f"{label}={cm}cm"
    if inches < 0 or inches * _PTS_PER_INCH > _MAX_DIMENSION_PT:
        raise ValueError(f"{label}={inches}in is out of plausible range")
    return float(inches) * _PTS_PER_INCH, f"{label}={inches}in"


async def word_live_set_page_layout(
    filename: str = None,
    section_index: int = 1,
    orientation: str = None,
    # Page size
    page_width_inches: float = None,
    page_height_inches: float = None,
    page_width_cm: float = None,
    page_height_cm: float = None,
    # Margins
    margin_top_inches: float = None,
    margin_bottom_inches: float = None,
    margin_left_inches: float = None,
    margin_right_inches: float = None,
    margin_top_cm: float = None,
    margin_bottom_cm: float = None,
    margin_left_cm: float = None,
    margin_right_cm: float = None,
    # Binding / header / footer distance
    gutter_cm: float = None,
    gutter_inches: float = None,
    header_distance_cm: float = None,
    header_distance_inches: float = None,
    footer_distance_cm: float = None,
    footer_distance_inches: float = None,
    # Section flags
    different_odd_and_even_pages: bool = None,
    different_first_page: bool = None,
) -> str:
    """Set page layout for a section in an open Word document.

    Units
    -----
    Each linear dimension accepts both ``_cm`` and ``_inches`` flavours; pass
    whichever matches your spec. If both are given for the same dimension,
    cm wins and a note is attached to the response. Internally all values
    are converted to Word's native unit (points: 72 pt/in, ~28.35 pt/cm).

    Supported properties
    --------------------
    * orientation: "portrait" or "landscape"
    * page width / height
    * top / bottom / left / right margins
    * gutter (binding margin) — useful for two-sided printing
    * header_distance — distance from page edge to header band
    * footer_distance — distance from page edge to footer band
    * different_odd_and_even_pages — section-level toggle for alternating
      odd/even page headers and footers
    * different_first_page — section-level toggle for a special first-page
      header/footer (e.g. cover page)

    Args:
        filename: Document name or path (None = active document).
        section_index: Section number (1-indexed, COM style). Default 1.
        ...: (each dimension can be given in either cm or inches)

    Returns:
        JSON with ``changes`` (human-readable log) and
        ``effective_pt`` (dict of every property actually applied in points,
        suitable for diff-against-spec verification).
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_set_page_layout
        # macOS path doesn't yet honor cm / gutter / header_distance / section flags;
        # forward only what it supports. (Issue tracked for the JXA backend.)
        return mac_set_page_layout(
            filename=filename, section_index=section_index, orientation=orientation,
            page_width=page_width_inches, page_height=page_height_inches,
            top_margin=margin_top_inches, bottom_margin=margin_bottom_inches,
            left_margin=margin_left_inches, right_margin=margin_right_inches,
        )

    if sys.platform != "win32":
        return json.dumps({"error": "Live layout tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        # Resolve every dimension once up-front so unit errors fail before
        # we open the undo_record (cleaner Ctrl+Z semantics).
        try:
            page_width_pt,  page_width_log  = _resolve_pt(page_width_cm,  page_width_inches,  "page_width")
            page_height_pt, page_height_log = _resolve_pt(page_height_cm, page_height_inches, "page_height")
            mtop_pt,    mtop_log    = _resolve_pt(margin_top_cm,    margin_top_inches,    "margin_top")
            mbottom_pt, mbottom_log = _resolve_pt(margin_bottom_cm, margin_bottom_inches, "margin_bottom")
            mleft_pt,   mleft_log   = _resolve_pt(margin_left_cm,   margin_left_inches,   "margin_left")
            mright_pt,  mright_log  = _resolve_pt(margin_right_cm,  margin_right_inches,  "margin_right")
            gutter_pt,  gutter_log  = _resolve_pt(gutter_cm,        gutter_inches,        "gutter")
            hdist_pt,   hdist_log   = _resolve_pt(header_distance_cm, header_distance_inches, "header_distance")
            fdist_pt,   fdist_log   = _resolve_pt(footer_distance_cm, footer_distance_inches, "footer_distance")
        except ValueError as ve:
            return json.dumps({"error": str(ve)})

        app = get_word_app()
        doc = find_document(app, filename)

        if section_index < 1 or section_index > doc.Sections.Count:
            return json.dumps({
                "error": f"Section {section_index} out of range (1-{doc.Sections.Count})"
            })

        effective_pt: dict[str, float | bool] = {}
        changes: list[str] = []

        with undo_record(app, "MCP: Set Page Layout"):
            ps = doc.Sections(section_index).PageSetup

            if orientation is not None:
                # wdOrientPortrait=0, wdOrientLandscape=1
                if orientation.lower() == "landscape":
                    ps.Orientation = 1
                    changes.append("orientation=landscape")
                    effective_pt["orientation"] = "landscape"
                elif orientation.lower() == "portrait":
                    ps.Orientation = 0
                    changes.append("orientation=portrait")
                    effective_pt["orientation"] = "portrait"

            def apply(prop: str, pt_value: float | None, log: str, key: str):
                if pt_value is not None:
                    setattr(ps, prop, pt_value)
                    changes.append(log)
                    effective_pt[key] = pt_value

            apply("PageWidth",       page_width_pt,  page_width_log,  "page_width_pt")
            apply("PageHeight",      page_height_pt, page_height_log, "page_height_pt")
            apply("TopMargin",       mtop_pt,        mtop_log,        "margin_top_pt")
            apply("BottomMargin",    mbottom_pt,     mbottom_log,     "margin_bottom_pt")
            apply("LeftMargin",      mleft_pt,       mleft_log,       "margin_left_pt")
            apply("RightMargin",     mright_pt,      mright_log,      "margin_right_pt")
            apply("Gutter",          gutter_pt,      gutter_log,      "gutter_pt")
            apply("HeaderDistance",  hdist_pt,       hdist_log,       "header_distance_pt")
            apply("FooterDistance",  fdist_pt,       fdist_log,       "footer_distance_pt")

            if different_odd_and_even_pages is not None:
                ps.OddAndEvenPagesHeaderFooter = bool(different_odd_and_even_pages)
                changes.append(f"different_odd_and_even_pages={different_odd_and_even_pages}")
                effective_pt["different_odd_and_even_pages"] = bool(different_odd_and_even_pages)

            if different_first_page is not None:
                ps.DifferentFirstPageHeaderFooter = bool(different_first_page)
                changes.append(f"different_first_page={different_first_page}")
                effective_pt["different_first_page"] = bool(different_first_page)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "section": section_index,
            "changes": changes,
            "effective_pt": effective_pt,
        })

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_add_header_footer(
    filename: str = None,
    section_index: int = 1,
    header_text: str = None,
    footer_text: str = None,
    header_alignment: str = "center",
    footer_alignment: str = "center",
) -> str:
    """Add header and/or footer text to a section in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        section_index: Section number (1-indexed).
        header_text: Text for the header. None = don't change.
        footer_text: Text for the footer. None = don't change.
        header_alignment: "left", "center", "right".
        footer_alignment: "left", "center", "right".

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_add_header_footer
        return mac_add_header_footer(filename=filename, section_index=section_index, header_text=header_text, footer_text=footer_text)

    if sys.platform != "win32":
        return json.dumps({"error": "Live layout tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if section_index < 1 or section_index > doc.Sections.Count:
            return json.dumps({
                "error": f"Section {section_index} out of range (1-{doc.Sections.Count})"
            })

        with undo_record(app, "MCP: Add Header/Footer"):
            # Alignment map: 0=left, 1=center, 2=right
            align_map = {"left": 0, "center": 1, "right": 2}
            added = []

            # wdHeaderFooterPrimary = 1
            section = doc.Sections(section_index)

            if header_text is not None:
                hdr = section.Headers(1)  # Primary header
                hdr.Range.Text = header_text
                hdr.Range.ParagraphFormat.Alignment = align_map.get(
                    header_alignment.lower(), 1
                )
                added.append("header")

            if footer_text is not None:
                ftr = section.Footers(1)  # Primary footer
                ftr.Range.Text = footer_text
                ftr.Range.ParagraphFormat.Alignment = align_map.get(
                    footer_alignment.lower(), 1
                )
                added.append("footer")

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "section": section_index,
            "added": added,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_add_page_numbers(
    filename: str = None,
    section_index: int = 1,
    position: str = "footer",
    alignment: str = "center",
    prefix: str = "",
    suffix: str = "",
    include_total: bool = False,
) -> str:
    """Add page numbers to header or footer in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        section_index: Section number (1-indexed).
        position: "header" or "footer".
        alignment: "left", "center", "right".
        prefix: Text before page number.
        suffix: Text after page number.
        include_total: If True, adds " / N" after page number.

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"error": "word_live_add_page_numbers is not yet implemented on macOS"})

    if sys.platform != "win32":
        return json.dumps({"error": "Live layout tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if section_index < 1 or section_index > doc.Sections.Count:
            return json.dumps({
                "error": f"Section {section_index} out of range (1-{doc.Sections.Count})"
            })

        with undo_record(app, "MCP: Add Page Numbers"):
            # PageNumberAlignment: 0=left, 1=center, 2=right
            align_map = {"left": 0, "center": 1, "right": 2}
            pn_alignment = align_map.get(alignment.lower(), 1)

            section = doc.Sections(section_index)
            # wdHeaderFooterPrimary = 1
            target = section.Headers(1) if position == "header" else section.Footers(1)

            # Add page numbers via PageNumbers collection
            target.PageNumbers.Add(PageNumberAlignment=pn_alignment)

            # Add prefix/suffix/total by editing the range
            if prefix or suffix or include_total:
                rng = target.Range
                existing_text = rng.Text

                # Build the text with field codes
                # Clear and rebuild
                rng.Delete()

                if prefix:
                    rng.InsertAfter(prefix)

                # Insert PAGE field
                # wdFieldPage = 33
                rng.Collapse(0)  # wdCollapseEnd
                app.Selection.GoTo(What=1, Name=str(section_index))  # navigate to section
                field_range = target.Range
                field_range.Collapse(0)
                doc.Fields.Add(Range=field_range, Type=33)  # wdFieldPage

                if include_total:
                    end_range = target.Range
                    end_range.Collapse(0)
                    end_range.InsertAfter(" / ")
                    end_range = target.Range
                    end_range.Collapse(0)
                    doc.Fields.Add(Range=end_range, Type=26)  # wdFieldNumPages

                if suffix:
                    end_range = target.Range
                    end_range.Collapse(0)
                    end_range.InsertAfter(suffix)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "section": section_index,
            "position": position,
            "alignment": alignment,
            "include_total": include_total,
        })

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_add_section_break(
    filename: str = None,
    break_type: str = "new_page",
) -> str:
    """Add a section break to an open Word document.

    Args:
        filename: Document name or path (None = active document).
        break_type: "new_page", "continuous", "even_page", "odd_page".

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_add_section_break
        return mac_add_section_break(filename=filename, break_type=break_type)

    if sys.platform != "win32":
        return json.dumps({"error": "Live layout tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        # wdSectionBreakNextPage=2, Continuous=3, EvenPage=4, OddPage=5
        type_map = {
            "new_page": 2,
            "continuous": 3,
            "even_page": 4,
            "odd_page": 5,
        }

        if break_type not in type_map:
            return json.dumps({
                "error": f"Invalid break_type: {break_type}. Use: {list(type_map.keys())}"
            })

        with undo_record(app, "MCP: Add Section Break"):
            # Insert at end of document
            end_pos = doc.Content.End - 1
            rng = doc.Range(end_pos, end_pos)
            rng.InsertBreak(Type=type_map[break_type])

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "break_type": break_type,
            "total_sections": doc.Sections.Count,
        })

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_set_paragraph_spacing(
    filename: str = None,
    paragraph_index: int = None,
    start_paragraph: int = None,
    end_paragraph: int = None,
    space_before_pt: float = None,
    space_after_pt: float = None,
    line_spacing: float = None,
    line_spacing_rule: str = None,
    keep_with_next: bool = None,
    keep_together: bool = None,
    alignment: str = None,
) -> str:
    """Set paragraph spacing and layout properties in an open Word document.

    Line-spacing interpretation (fixed in 1.7.2 — previous releases required
    you to pre-multiply by 12, which silently produced 1.5-pt fixed spacing
    when callers passed 1.5):

      * ``line_spacing_rule="multiple"`` — pass a multiplier (1.0, 1.15, 1.5,
        2.0…). The tool converts to Word COM's required points value as
        ``multiplier × 12``. Values > 10 are treated as already-points
        (backward compat for callers who pre-multiplied).
      * ``line_spacing_rule="exactly"`` / ``"at_least"`` — value is in
        points (e.g. 18 means literally 18 pt).
      * ``line_spacing_rule="single"`` / ``"1.5_lines"`` / ``"double"`` —
        Word computes the points value itself; any ``line_spacing`` you
        pass is ignored.
      * ``line_spacing_rule=None`` — inferred: values ≤ 10 use ``"multiple"``
        (×12 conversion), values > 10 use ``"exactly"`` as-is.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index: Single paragraph (1-indexed). Ignored if start/end given.
        start_paragraph: Start of range (1-indexed, inclusive).
        end_paragraph: End of range (1-indexed, inclusive).
        space_before_pt: Space before paragraph in points.
        space_after_pt: Space after paragraph in points.
        line_spacing: Multiplier or points (see interpretation block above).
        line_spacing_rule: "single", "1.5_lines", "double", "at_least",
            "exactly", or "multiple". Defaults to inferred.
        keep_with_next: Keep paragraph with next paragraph on same page (True/False).
        keep_together: Keep all lines of paragraph on same page (True/False).
        alignment: Paragraph alignment - "left", "center", "right", "justify".

    Returns:
        JSON with paragraphs_affected and the ``effective_line_spacing_pt`` /
        ``effective_line_spacing_rule`` we actually applied (useful for
        diagnosing surprises). On macOS the JXA backend does not yet honor
        ``line_spacing_rule``; only ``line_spacing`` (as points) is applied.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_set_paragraph_spacing
        return mac_set_paragraph_spacing(filename=filename, paragraph_index=paragraph_index, start_paragraph=start_paragraph, end_paragraph=end_paragraph, space_before=space_before_pt, space_after=space_after_pt, line_spacing=line_spacing, keep_with_next=keep_with_next, keep_together=keep_together, alignment=alignment)

    if sys.platform != "win32":
        return json.dumps({"error": "Live layout tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.Paragraphs.Count

        # wdLineSpacing rules
        rule_map = {
            "single": 0,
            "1.5_lines": 1,
            "double": 2,
            "at_least": 3,
            "exactly": 4,
            "multiple": 5,
        }
        # Word's preset rules ignore the numeric LineSpacing value entirely.
        PRESET_RULES = {"single", "1.5_lines", "double"}

        # ---- Normalize line_spacing against rule. -----------------------
        # Word COM Paragraph.LineSpacing is ALWAYS in points. For the
        # "multiple" rule Word still wants points; the conversion is
        # multiplier × 12 (12 being Word's single-line baseline). Passing
        # a raw 1.5 with rule="multiple" gives 1.5-pt fixed line height —
        # text glues together. We accept the natural multiplier form here
        # and convert for the caller.
        effective_rule = line_spacing_rule
        effective_spacing = line_spacing
        if line_spacing is not None:
            if effective_rule in PRESET_RULES:
                # Preset rules compute their own value; drop user's number
                # to make behavior predictable.
                effective_spacing = None
            elif effective_rule == "multiple":
                if line_spacing <= 10:
                    effective_spacing = float(line_spacing) * 12.0
                # else: assume caller already pre-multiplied (>10 pts)
            elif effective_rule in ("exactly", "at_least"):
                effective_spacing = float(line_spacing)  # already points
            elif effective_rule is None:
                # Infer rule from magnitude.
                if line_spacing <= 10:
                    effective_rule = "multiple"
                    effective_spacing = float(line_spacing) * 12.0
                else:
                    effective_rule = "exactly"
                    effective_spacing = float(line_spacing)

        # Determine range of paragraphs (1-indexed)
        if start_paragraph is not None and end_paragraph is not None:
            indices = range(max(1, start_paragraph), min(end_paragraph + 1, total + 1))
        elif paragraph_index is not None:
            if paragraph_index < 1 or paragraph_index > total:
                return json.dumps({
                    "error": f"paragraph_index {paragraph_index} out of range (1-{total})"
                })
            indices = [paragraph_index]
        else:
            indices = range(1, total + 1)

        with undo_record(app, "MCP: Set Paragraph Spacing"):
            count = 0
            for i in indices:
                pf = doc.Paragraphs(i).Format
                if space_before_pt is not None:
                    pf.SpaceBefore = space_before_pt
                if space_after_pt is not None:
                    pf.SpaceAfter = space_after_pt
                # Rule must be set BEFORE LineSpacing or Word may reinterpret.
                if effective_rule is not None and effective_rule in rule_map:
                    pf.LineSpacingRule = rule_map[effective_rule]
                if effective_spacing is not None:
                    pf.LineSpacing = effective_spacing
                if keep_with_next is not None:
                    pf.KeepWithNext = keep_with_next
                if keep_together is not None:
                    pf.KeepTogether = keep_together
                if alignment is not None:
                    align_map = {"left": 0, "center": 1, "right": 2, "justify": 3}
                    if alignment in align_map:
                        pf.Alignment = align_map[alignment]
                count += 1

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraphs_affected": count,
            "effective_line_spacing_rule": effective_rule,
            "effective_line_spacing_pt": effective_spacing,
        })

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_add_bookmark(
    filename: str = None,
    paragraph_index: int = 1,
    bookmark_name: str = "",
) -> str:
    """Add a named bookmark at a paragraph in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index: Paragraph to bookmark (1-indexed).
        bookmark_name: Bookmark name (alphanumeric + underscore, no spaces).

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_add_bookmark
        return mac_add_bookmark(filename=filename, paragraph_index=paragraph_index, bookmark_name=bookmark_name)

    if sys.platform != "win32":
        return json.dumps({"error": "Live layout tools are only available on Windows"})

    if not bookmark_name:
        return json.dumps({"error": "bookmark_name is required"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
            return json.dumps({
                "error": f"paragraph_index {paragraph_index} out of range (1-{doc.Paragraphs.Count})"
            })

        with undo_record(app, "MCP: Add Bookmark"):
            rng = doc.Paragraphs(paragraph_index).Range
            doc.Bookmarks.Add(bookmark_name, rng)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "bookmark_name": bookmark_name,
            "paragraph_index": paragraph_index,
        })

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_add_watermark(
    filename: str = None,
    text: str = "TASLAK",
    font_size: int = 72,
    font_color: str = "C0C0C0",
    rotation: int = -45,
    section_index: int = 1,
) -> str:
    """Add a diagonal text watermark to an open Word document.

    Args:
        filename: Document name or path (None = active document).
        text: Watermark text (e.g. "TASLAK", "DRAFT", "GİZLİ").
        font_size: Font size in points.
        font_color: Hex color without # (e.g. "C0C0C0").
        rotation: Rotation angle in degrees (e.g. -45).
        section_index: Section number (1-indexed).

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"error": "word_live_add_watermark is not yet implemented on macOS"})

    if sys.platform != "win32":
        return json.dumps({"error": "Live layout tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if section_index < 1 or section_index > doc.Sections.Count:
            return json.dumps({
                "error": f"Section {section_index} out of range (1-{doc.Sections.Count})"
            })

        with undo_record(app, "MCP: Add Watermark"):
            section = doc.Sections(section_index)
            header = section.Headers(1)  # wdHeaderFooterPrimary

            # Parse color
            c = font_color.lstrip("#")
            r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
            rgb_color = r + (g << 8) + (b << 16)

            # AddTextEffect(PresetTextEffect, Text, FontName, FontSize,
            #               FontBold, FontItalic, Left, Top)
            # COM requires positional args
            shape = header.Shapes.AddTextEffect(
                0, text, "Calibri", font_size, False, False, 0, 0
            )

            # Configure shape
            shape.Fill.ForeColor.RGB = rgb_color
            shape.Fill.Transparency = 0.5
            shape.Line.Visible = False  # msoFalse
            shape.Rotation = rotation
            shape.LockAspectRatio = False

            # Position relative to page center
            # msoRelativeHorizontalPositionMargin = 0
            # msoRelativeVerticalPositionMargin = 0
            shape.RelativeHorizontalPosition = 0
            shape.RelativeVerticalPosition = 0
            shape.Left = -999995  # wdShapeCenter (magic value for centering)
            shape.Top = -999995  # wdShapeCenter

            # Send behind text
            shape.WrapFormat.Type = 3  # wdWrapBehind
            shape.WrapFormat.AllowOverlap = True

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "text": text,
            "font_size": font_size,
            "color": font_color,
            "rotation": rotation,
            "section": section_index,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})
