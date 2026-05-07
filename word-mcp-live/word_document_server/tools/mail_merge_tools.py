"""Mail Merge tools for open Word documents via COM.

Workflow:
  1. word_live_setup_mail_merge   — open the document and connect a data source
  2. word_live_list_merge_fields  — inspect available data columns
  3. word_live_insert_merge_field — add «FieldName» placeholders to the template
  4. word_live_preview_merge      — preview a specific record in Word
  5. word_live_execute_merge      — produce the final merged output
"""

import json
import os
import sys


async def word_live_setup_mail_merge(
    filename: str = None,
    data_source_path: str = None,
    merge_type: str = "letters",
    header_row: bool = True,
) -> str:
    """Connect a Word document to a mail merge data source.

    Supported data sources: Excel (.xlsx/.xls), CSV (.csv),
    Access (.mdb/.accdb), plain text (.txt with tab or comma delimiters).

    Args:
        filename: Template document name or path (None = active document).
        data_source_path: Path to the data file (Excel, CSV, etc.).
        merge_type: "letters" (default), "email", "envelopes", "labels", "catalog".
        header_row: First row is a header (column names). Default True.

    Returns:
        JSON with data source info and field names.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Mail merge tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        MERGE_TYPE_MAP = {
            "letters":   0,   # wdMailMergeFormLetters
            "email":     1,   # wdMailMergeEMail
            "envelopes": 2,   # wdMailMergeEnvelopes
            "labels":    3,   # wdMailMergeLabels
            "catalog":   4,   # wdMailMergeCatalog
        }

        app = get_word_app()
        doc = find_document(app, filename)
        mm = doc.MailMerge

        mm.MainDocumentType = MERGE_TYPE_MAP.get(merge_type, 0)

        if data_source_path:
            data_source_path = os.path.abspath(data_source_path)
            if not os.path.exists(data_source_path):
                return json.dumps({"error": f"Data source not found: {data_source_path}"})

            # Open the data source
            mm.OpenDataSource(
                Name=data_source_path,
                ConfirmConversions=False,
                ReadOnly=True,
                LinkToSource=True,
                AddToRecentFiles=False,
                PasswordDocument="",
                PasswordTemplate="",
                Revert=False,
                Format=-1,        # wdOpenFormatAuto
                Connection="",
                SQLStatement="",
                SQLStatement1="",
                SubType=0,
            )

        # Read field names
        field_names = []
        try:
            ds = mm.DataSource
            for i in range(1, ds.FieldNames.Count + 1):
                field_names.append(ds.FieldNames(i).Name)
        except Exception:
            pass

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "data_source": data_source_path or "already connected",
            "merge_type": merge_type,
            "record_count": mm.DataSource.RecordCount if data_source_path else None,
            "field_names": field_names,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_list_merge_fields(filename: str = None) -> str:
    """List all data field names available from the connected mail merge data source.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with field names and record count.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Mail merge tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)
        mm = doc.MailMerge

        ds = mm.DataSource
        field_names = [ds.FieldNames(i).Name for i in range(1, ds.FieldNames.Count + 1)]

        # Also list merge fields already in the document
        doc_fields = []
        for i in range(1, mm.Fields.Count + 1):
            try:
                doc_fields.append(mm.Fields(i).Name)
            except Exception:
                pass

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "data_field_names": field_names,
            "record_count": ds.RecordCount,
            "merge_fields_in_document": doc_fields,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_insert_merge_field(
    filename: str = None,
    field_name: str = "",
    paragraph_index: int = None,
    position: str = "end",
) -> str:
    """Insert a «MergeField» placeholder into a mail merge template.

    Args:
        filename: Document name or path (None = active document).
        field_name: Data source column name to insert (e.g. "FirstName", "Address").
        paragraph_index: Target paragraph (1-indexed). None = end of document.
        position: "end" (default) or "start" within the paragraph.

    Returns:
        JSON confirming insertion.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Mail merge tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        with undo_record(app, f"MCP: Insert Merge Field {field_name}"):
            if paragraph_index is not None:
                total = doc.Paragraphs.Count
                if paragraph_index < 1 or paragraph_index > total:
                    return json.dumps({"error": f"paragraph_index {paragraph_index} out of range"})
                p = doc.Paragraphs(paragraph_index).Range
                ins_pos = p.Start if position == "start" else p.End - 1
                rng = doc.Range(ins_pos, ins_pos)
            else:
                end = doc.Content.End - 1
                rng = doc.Range(end, end)

            doc.MailMerge.Fields.Add(Range=rng, Name=field_name)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "field_name": field_name,
            "total_merge_fields": doc.MailMerge.Fields.Count,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_preview_merge(
    filename: str = None,
    record_index: int = 1,
) -> str:
    """Preview a specific data record in the mail merge template in Word.

    Fills the merge fields with the data from the specified record so you can
    see how the merged result will look. Call again with record_index=0 to
    turn off preview and return to the template view.

    Args:
        filename: Document name or path (None = active document).
        record_index: Data record to preview (1-based). 0 = turn off preview.

    Returns:
        JSON confirming the preview state.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Mail merge tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)
        mm = doc.MailMerge

        if record_index <= 0:
            mm.ViewMailMergeFieldCodes = True   # show field codes (template view)
            return json.dumps({"success": True, "mode": "template_view"})

        mm.DataSource.ActiveRecord = record_index
        mm.ViewMailMergeFieldCodes = False   # show data (preview mode)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "previewing_record": record_index,
            "total_records": mm.DataSource.RecordCount,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_execute_merge(
    filename: str = None,
    output_path: str = None,
    records: str = "all",
    from_record: int = 1,
    to_record: int = None,
    merge_to: str = "new_document",
    email_address_field: str = None,
    email_subject: str = None,
) -> str:
    """Execute the mail merge and produce output documents.

    Args:
        filename: Template document name or path (None = active document).
        output_path: Where to save the merged output file. None = open in Word
            as a new unsaved document (merge_to must be "new_document").
        records: "all" (default), "from_to" (use from_record/to_record), or
                 "current" (only the currently previewed record).
        from_record: First record to merge (used when records="from_to").
        to_record: Last record to merge (used when records="from_to"). None = last.
        merge_to: "new_document" (default), "printer", or "email".
        email_address_field: Field name containing email addresses (merge_to="email").
        email_subject: Email subject line (merge_to="email").

    Returns:
        JSON with the path of the merged output document (or status for printer/email).
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Mail merge tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        RECORDS_MAP = {"all": 1, "from_to": 3, "current": 5}
        # wdSendToNewDocument=0, wdSendToPrinter=1, wdSendToEmail=2
        DEST_MAP = {"new_document": 0, "printer": 1, "email": 2}

        app = get_word_app()
        doc = find_document(app, filename)
        mm = doc.MailMerge

        # Set record range
        record_flag = RECORDS_MAP.get(records, 1)
        if records == "from_to":
            mm.DataSource.FirstRecord = from_record
            mm.DataSource.LastRecord = to_record or mm.DataSource.RecordCount
        elif records == "all":
            # wdDefaultFirstRecord/LastRecord
            mm.DataSource.FirstRecord = 1
            mm.DataSource.LastRecord = -16     # wdLastRecord constant

        dest = DEST_MAP.get(merge_to, 0)

        if merge_to == "email":
            if not email_address_field:
                return json.dumps({"error": "email_address_field is required for merge_to='email'"})
            mm.MailAddressFieldName = email_address_field
            if email_subject:
                mm.MailSubject = email_subject

        mm.Destination = dest
        mm.Execute(Pause=False)

        # If new document, save to output_path
        result_info = {"success": True, "document": doc.Name, "merge_to": merge_to}

        if merge_to == "new_document":
            merged_doc = app.ActiveDocument
            if output_path:
                output_path = os.path.abspath(output_path)
                merged_doc.SaveAs2(output_path, FileFormat=16)
                result_info["output_path"] = output_path
            else:
                result_info["merged_document"] = merged_doc.Name
                result_info["note"] = "Merged document is open in Word but not saved."

        return json.dumps(result_info)
    except Exception as e:
        return json.dumps({"error": str(e)})
