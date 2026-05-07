"""Session management tools — open, close, create, list Word documents via COM."""

import json
import os
import sys


async def word_open_document(
    path: str,
    visible: bool = True,
    read_only: bool = False,
    password: str = None,
) -> str:
    """Open a Word document (.docx / .doc) in the live Word application.

    Launches Word automatically if it is not already running.

    Args:
        path: Absolute or relative path to the document file.
        visible: Make the Word window visible after opening (default True).
        read_only: Open in read-only mode (default False).
        password: Password to open a protected document.

    Returns:
        JSON with document name, path, and page/paragraph counts.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Session tools are only available on Windows"})
    try:
        from word_document_server.core.word_com import get_word_app

        path = os.path.abspath(path)
        if not os.path.exists(path):
            return json.dumps({"error": f"File not found: {path}"})

        app = get_word_app(visible=visible)
        kwargs = {"FileName": path, "ReadOnly": read_only}
        if password:
            kwargs["PasswordDocument"] = password
        doc = app.Documents.Open(**kwargs)
        if visible:
            app.Visible = True
            app.Activate()

        return json.dumps({
            "success": True,
            "name": doc.Name,
            "path": doc.FullName,
            "pages": doc.ComputeStatistics(2),   # wdStatisticPages=2
            "paragraphs": doc.Paragraphs.Count,
            "read_only": read_only,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_new_document(
    template: str = None,
    visible: bool = True,
    title: str = None,
) -> str:
    """Create a new blank Word document (optionally from a template).

    Args:
        template: Path to a .dotx/.dotm template file, or None for blank.
        visible: Make the Word window visible (default True).
        title: Set the document title in its core properties.

    Returns:
        JSON with the new document name and temporary path.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Session tools are only available on Windows"})
    try:
        from word_document_server.core.word_com import get_word_app

        app = get_word_app(visible=visible)
        if template:
            template = os.path.abspath(template)
            doc = app.Documents.Add(Template=template)
        else:
            doc = app.Documents.Add()

        if visible:
            app.Visible = True
            app.Activate()

        if title:
            try:
                doc.BuiltInDocumentProperties("Title").Value = title
            except Exception:
                pass

        return json.dumps({
            "success": True,
            "name": doc.Name,
            "path": doc.FullName,
            "template": template or "Normal",
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_close_document(
    filename: str = None,
    save: bool = True,
    save_path: str = None,
) -> str:
    """Close an open Word document.

    Args:
        filename: Document name or path (None = active document).
        save: Save before closing (default True). False discards changes.
        save_path: Save to a different path before closing (Save As).

    Returns:
        JSON confirming closure.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Session tools are only available on Windows"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app(auto_launch=False)
        doc = find_document(app, filename)
        name = doc.Name
        path = doc.FullName

        if save_path:
            save_path = os.path.abspath(save_path)
            doc.SaveAs2(save_path, FileFormat=16)
        elif save:
            if doc.Path:
                doc.Save()
            else:
                return json.dumps({
                    "error": "Document has never been saved. Provide save_path to save it first."
                })

        doc.Close(SaveChanges=False)
        return json.dumps({"success": True, "closed": name, "path": path, "saved": save})
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_save_document(
    filename: str = None,
    save_path: str = None,
    format: str = "docx",
) -> str:
    """Save an open Word document, optionally to a new path.

    Args:
        filename: Document name or path (None = active document).
        save_path: New save path (Save As). None = save in place.
        format: Output format — "docx" (default), "doc", "pdf", "txt", "rtf".

    Returns:
        JSON confirming save and final path.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Session tools are only available on Windows"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        FORMAT_MAP = {
            "docx": 16,   # wdFormatXMLDocument
            "doc":  0,    # wdFormatDocument
            "pdf":  17,   # wdFormatPDF
            "txt":  2,    # wdFormatText
            "rtf":  6,    # wdFormatRTF
        }
        fmt = FORMAT_MAP.get(format.lower(), 16)

        app = get_word_app(auto_launch=False)
        doc = find_document(app, filename)

        if save_path:
            save_path = os.path.abspath(save_path)
            doc.SaveAs2(save_path, FileFormat=fmt)
            final_path = save_path
        else:
            if not doc.Path:
                return json.dumps({
                    "error": "Document has no path yet. Provide save_path."
                })
            doc.Save()
            final_path = doc.FullName

        return json.dumps({
            "success": True,
            "name": doc.Name,
            "path": final_path,
            "format": format,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_list_sessions() -> str:
    """List all documents currently open in Word.

    Returns:
        JSON array of open documents with name, path, and modification state.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Session tools are only available on Windows"})
    try:
        from word_document_server.core.word_com import get_word_app

        app = get_word_app(auto_launch=False)
        docs = []
        for i in range(1, app.Documents.Count + 1):
            d = app.Documents(i)
            is_active = False
            try:
                is_active = app.ActiveDocument.Name == d.Name
            except Exception:
                pass
            docs.append({
                "index": i,
                "name": d.Name,
                "path": d.FullName,
                "modified": d.Saved is False,
                "read_only": d.ReadOnly,
                "active": is_active,
                "pages": d.ComputeStatistics(2),
                "paragraphs": d.Paragraphs.Count,
            })
        return json.dumps({"success": True, "count": len(docs), "documents": docs})
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_activate_document(filename: str) -> str:
    """Bring a specific open document to the foreground in Word.

    Args:
        filename: Document name or path.

    Returns:
        JSON confirming activation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Session tools are only available on Windows"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app(auto_launch=False)
        doc = find_document(app, filename)
        doc.Activate()
        app.Visible = True
        return json.dumps({"success": True, "activated": doc.Name})
    except Exception as e:
        return json.dumps({"error": str(e)})
