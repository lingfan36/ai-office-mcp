"""Field code management and TOC update tools for open Word documents via COM."""

import json
import sys

# Common wdFieldType constants
FIELD_TYPES = {
    "date":           31,   # wdFieldDate
    "time":           32,   # wdFieldTime
    "page":           33,   # wdFieldPage
    "num_pages":      26,   # wdFieldNumPages
    "filename":       29,   # wdFieldFileName
    "author":         17,   # wdFieldAuthor
    "title":          15,   # wdFieldTitle
    "subject":        14,   # wdFieldSubject
    "toc":            13,   # wdFieldTOC
    "seq":            86,   # wdFieldSequence
    "hyperlink":      88,   # wdFieldHyperlink
    "ref":            3,    # wdFieldRef (cross-reference)
    "if":             7,    # wdFieldIf
    "merge_field":    59,   # wdFieldMergeField
    "custom":         -1,   # use raw field_code string
}


async def word_live_list_fields(filename: str = None) -> str:
    """List all field codes in an open Word document.

    Returns:
        JSON array with field index, type, code, and current result text.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Field tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        fields = []
        for i in range(1, doc.Fields.Count + 1):
            f = doc.Fields(i)
            try:
                code = f.Code.Text.strip()
            except Exception:
                code = ""
            try:
                result = f.Result.Text.strip()
            except Exception:
                result = ""
            fields.append({
                "index": i,
                "type": f.Type,
                "code": code,
                "result": result,
                "locked": f.Locked,
            })

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "count": len(fields),
            "fields": fields,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_insert_field(
    filename: str = None,
    field_type: str = "date",
    paragraph_index: int = None,
    position: str = "end",
    field_code: str = None,
    preserve_formatting: bool = True,
) -> str:
    """Insert a field code into an open Word document.

    Common field types:
      "date"       → current date   (e.g. code: "DATE \\@ \"yyyy-MM-dd\"")
      "time"       → current time
      "page"       → current page number
      "num_pages"  → total page count
      "filename"   → file name
      "author"     → document author
      "title"      → document title
      "seq"        → auto-incrementing sequence (e.g. for figure numbers)
      "custom"     → use field_code directly (e.g. "DOCPROPERTY Company")

    Args:
        filename: Document name or path (None = active document).
        field_type: Field type key (see list above, default "date").
        paragraph_index: Insert into this paragraph (1-indexed). None = end of document.
        position: Where within the paragraph — "end" (default) or "start".
        field_code: Raw field code string for field_type="custom" or to override defaults.
            Examples: "DATE \\@ \"MMMM d, yyyy\"", "DOCPROPERTY Company", "SEQ Figure"
        preserve_formatting: Keep field result formatting on update (default True).

    Returns:
        JSON with field index and initial result.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Field tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        type_id = FIELD_TYPES.get(field_type, -1)

        with undo_record(app, f"MCP: Insert Field {field_type}"):
            if paragraph_index is not None:
                total = doc.Paragraphs.Count
                if paragraph_index < 1 or paragraph_index > total:
                    return json.dumps({"error": f"paragraph_index {paragraph_index} out of range"})
                p_range = doc.Paragraphs(paragraph_index).Range
                if position == "start":
                    ins_pos = p_range.Start
                else:
                    ins_pos = p_range.End - 1
                rng = doc.Range(ins_pos, ins_pos)
            else:
                end = doc.Content.End - 1
                rng = doc.Range(end, end)

            if type_id == -1 or field_code:
                # Insert via raw code
                code = field_code or f" {field_type.upper()} "
                f = doc.Fields.Add(
                    Range=rng,
                    Type=1,           # wdFieldEmpty
                    Text=code,
                    PreserveFormatting=preserve_formatting,
                )
            else:
                f = doc.Fields.Add(
                    Range=rng,
                    Type=type_id,
                    PreserveFormatting=preserve_formatting,
                )

            try:
                result_text = f.Result.Text.strip()
            except Exception:
                result_text = ""

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "field_index": doc.Fields.Count,
            "field_type": field_type,
            "result": result_text,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_update_fields(
    filename: str = None,
    field_index: int = None,
) -> str:
    """Update (recalculate) fields in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        field_index: Update only this field (1-based). None = update ALL fields.

    Returns:
        JSON confirming update count.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Field tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        if field_index is not None:
            total = doc.Fields.Count
            if field_index < 1 or field_index > total:
                return json.dumps({"error": f"field_index {field_index} out of range (1-{total})"})
            doc.Fields(field_index).Update()
            count = 1
        else:
            doc.Fields.Update()
            count = doc.Fields.Count

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "fields_updated": count,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_delete_field(
    filename: str = None,
    field_index: int = 1,
    replace_with_result: bool = True,
) -> str:
    """Delete a field from an open Word document.

    Args:
        filename: Document name or path (None = active document).
        field_index: 1-based field index (use word_live_list_fields).
        replace_with_result: Keep the field's current text after deletion (default True).
            False = delete the field AND its text.

    Returns:
        JSON confirming deletion.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Field tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.Fields.Count
        if field_index < 1 or field_index > total:
            return json.dumps({"error": f"field_index {field_index} out of range (1-{total})"})

        field = doc.Fields(field_index)
        code = ""
        try:
            code = field.Code.Text.strip()
        except Exception:
            pass

        with undo_record(app, "MCP: Delete Field"):
            if replace_with_result:
                field.Unlink()   # converts field to plain text with its current result
            else:
                field.Delete()

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "deleted_field_index": field_index,
            "code": code,
            "kept_result_text": replace_with_result,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_update_toc(
    filename: str = None,
    toc_index: int = 1,
    update_page_numbers_only: bool = False,
) -> str:
    """Update (refresh) a Table of Contents in an open Word document.

    Call this after adding, removing, or renaming headings to keep the TOC current.

    Args:
        filename: Document name or path (None = active document).
        toc_index: 1-based TOC index (default 1 = first TOC).
        update_page_numbers_only: True = refresh page numbers only (faster).
            False = rebuild the entire TOC including entry text (default).

    Returns:
        JSON confirming update.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Field tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.TablesOfContents.Count
        if total == 0:
            return json.dumps({
                "error": "No Table of Contents found in this document. "
                         "Add one first with add_table_of_contents."
            })
        if toc_index < 1 or toc_index > total:
            return json.dumps({"error": f"toc_index {toc_index} out of range (1-{total})"})

        toc = doc.TablesOfContents(toc_index)
        # wdUpdateEntireTable=1, wdUpdatePageNumbersOnly=2
        toc.Update() if not update_page_numbers_only else toc.UpdatePageNumbers()

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "toc_index": toc_index,
            "updated": "page_numbers_only" if update_page_numbers_only else "full",
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_update_all_tocs(filename: str = None) -> str:
    """Update all Tables of Contents in an open Word document.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with count of TOCs updated.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Field tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        count = doc.TablesOfContents.Count
        for i in range(1, count + 1):
            doc.TablesOfContents(i).Update()

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "tocs_updated": count,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})
