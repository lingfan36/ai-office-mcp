"""
Footnote and endnote tools for Word Document Server.

These tools handle footnote and endnote functionality,
including adding, customizing, and converting between them.

Canonical implementations
-------------------------
* **Add footnote** — `add_footnote_robust_tool` (Dict return, full validation).
* **Delete footnote** — `delete_footnote_robust_tool` (Dict return, orphan cleanup).
* **Validate** — `validate_footnotes_tool`.

The following entry points are kept for backward compatibility and now
forward to the canonical implementations above:

    add_footnote_to_document       →  add_footnote_robust (paragraph_index)
    add_footnote_enhanced          →  add_footnote_robust (paragraph_index)
    add_footnote_after_text        →  add_footnote_robust (position="after")
    add_footnote_before_text       →  add_footnote_robust (position="before")
    delete_footnote_from_document  →  delete_footnote_robust

The string-returning variants are convenient for legacy clients but return
less information. New code should prefer the `*_robust_tool` versions.
"""
import os
from typing import Optional, Dict, Any
from docx import Document
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE

from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension, get_file_lock
from word_document_server.core.footnotes import (
    find_footnote_references,
    get_format_symbols,
    customize_footnote_formatting,
    add_footnote_robust,
    add_endnote_robust,
    convert_footnotes_to_endnotes_robust,
    delete_footnote_robust,
    validate_document_footnotes,
    add_footnote_at_paragraph_end  # Compatibility function
)


async def add_footnote_to_document(filename: str, paragraph_index: int, footnote_text: str) -> str:
    """[DEPRECATED — prefer `add_footnote_robust` for richer return + validation.]

    Add a footnote to a specific paragraph in a Word document.

    Now delegates to the robust implementation (was previously a separate
    legacy path that fell back to inserting a literal "¹" character if
    python-docx's add_footnote was unavailable — that fallback produced
    invalid documents and is removed).

    Args:
        filename: Path to the Word document
        paragraph_index: Index of the paragraph to add footnote to (0-based)
        footnote_text: Text content of the footnote
    """
    filename = ensure_docx_extension(filename)

    # Validate paragraph_index here so we keep the legacy error string
    # callers may depend on.
    try:
        paragraph_index = int(paragraph_index)
    except (ValueError, TypeError):
        return "Invalid parameter: paragraph_index must be an integer"

    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}. Consider creating a copy first."

    try:
        async with get_file_lock(filename):
            success, message, _details = add_footnote_robust(
                filename=filename,
                paragraph_index=paragraph_index,
                footnote_text=footnote_text,
                validate_location=True,
            )
        return message
    except Exception as e:
        return f"Failed to add footnote: {str(e)}"


async def add_endnote_to_document(filename: str, paragraph_index: int, endnote_text: str) -> str:
    """Add a native Word endnote to a specific paragraph.

    Writes a real ``<w:endnoteReference>`` into the paragraph and a matching
    ``<w:endnote>`` into ``word/endnotes.xml`` (creating the part, content-type
    override, relationship, and EndnoteReference/EndnoteText styles if they
    don't yet exist). Word recognises the result as a native endnote — visible
    in View > Endnotes, accept/reject via the Review pane, renumbered
    automatically when others are added.

    (Previous versions appended a "†" superscript and a fake "Endnotes:"
    heading paragraph at the document tail. That was not a real endnote;
    Word would not number, group, or convert it.)

    Args:
        filename: Path to the Word document
        paragraph_index: Index of the paragraph to add endnote to (0-based)
        endnote_text: Text content of the endnote
    """
    filename = ensure_docx_extension(filename)

    try:
        paragraph_index = int(paragraph_index)
    except (ValueError, TypeError):
        return "Invalid parameter: paragraph_index must be an integer"

    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}. Consider creating a copy first."

    try:
        async with get_file_lock(filename):
            success, message, _details = add_endnote_robust(
                filename=filename,
                paragraph_index=paragraph_index,
                endnote_text=endnote_text,
                validate_location=True,
            )
        return message
    except Exception as e:
        return f"Failed to add endnote: {str(e)}"


async def convert_footnotes_to_endnotes_in_document(filename: str) -> str:
    """Convert every native footnote in the document to a native endnote.

    Walks ``word/footnotes.xml``, deep-copies each non-separator footnote
    into ``word/endnotes.xml`` (renaming tags and styles), removes the
    originals, then rewrites every ``<w:footnoteReference>`` in the body to
    ``<w:endnoteReference>`` under the new id, flipping the surrounding
    rStyle from FootnoteReference to EndnoteReference. Endnote part,
    content-type override, relationship, and EndnoteReference/EndnoteText
    styles are created if missing.

    (Previous versions scanned for "¹²³..." superscript runs and appended
    a fake "Endnotes:" heading — that produced documents Word did not
    recognise as having endnotes.)

    Args:
        filename: Path to the Word document
    """
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}. Consider creating a copy first."

    try:
        async with get_file_lock(filename):
            success, message, _details = convert_footnotes_to_endnotes_robust(filename=filename)
        return message
    except Exception as e:
        return f"Failed to convert footnotes to endnotes: {str(e)}"


async def add_footnote_after_text(filename: str, search_text: str, footnote_text: str, 
                                 output_filename: Optional[str] = None) -> str:
    """Add a footnote after specific text in a Word document with proper formatting.
    
    This enhanced function ensures proper superscript formatting by managing styles at the XML level.
    
    Args:
        filename: Path to the Word document
        search_text: Text to search for (footnote will be added after this text)
        footnote_text: Content of the footnote
        output_filename: Optional output filename (if None, modifies in place)
    """
    filename = ensure_docx_extension(filename)
    
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    
    # Check if file is writeable
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}. Consider creating a copy first."
    
    try:
        async with get_file_lock(filename):
            success, message, details = add_footnote_robust(
                filename=filename,
                search_text=search_text,
                footnote_text=footnote_text,
                output_filename=output_filename,
                position="after",
                validate_location=True
            )
        return message
    except Exception as e:
        return f"Failed to add footnote: {str(e)}"


async def add_footnote_before_text(filename: str, search_text: str, footnote_text: str, 
                                  output_filename: Optional[str] = None) -> str:
    """Add a footnote before specific text in a Word document with proper formatting.
    
    This enhanced function ensures proper superscript formatting by managing styles at the XML level.
    
    Args:
        filename: Path to the Word document
        search_text: Text to search for (footnote will be added before this text)
        footnote_text: Content of the footnote
        output_filename: Optional output filename (if None, modifies in place)
    """
    filename = ensure_docx_extension(filename)
    
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    
    # Check if file is writeable
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}. Consider creating a copy first."
    
    try:
        async with get_file_lock(filename):
            success, message, details = add_footnote_robust(
                filename=filename,
                search_text=search_text,
                footnote_text=footnote_text,
                output_filename=output_filename,
                position="before",
                validate_location=True
            )
        return message
    except Exception as e:
        return f"Failed to add footnote: {str(e)}"


async def add_footnote_enhanced(filename: str, paragraph_index: int, footnote_text: str,
                               output_filename: Optional[str] = None) -> str:
    """[DEPRECATED — same behavior as `add_footnote_robust`; prefer that instead.]

    Originally an "enhanced" upgrade over `add_footnote_to_document`; now both
    delegate to `add_footnote_robust`, so this entry point exists only for
    backward compatibility.

    Args:
        filename: Path to the Word document
        paragraph_index: Index of the paragraph to add footnote to (0-based)
        footnote_text: Text content of the footnote
        output_filename: Optional output filename (if None, modifies in place)
    """
    filename = ensure_docx_extension(filename)
    
    # Ensure paragraph_index is an integer
    try:
        paragraph_index = int(paragraph_index)
    except (ValueError, TypeError):
        return "Invalid parameter: paragraph_index must be an integer"
    
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    
    # Check if file is writeable
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}. Consider creating a copy first."
    
    try:
        async with get_file_lock(filename):
            success, message, details = add_footnote_robust(
                filename=filename,
                paragraph_index=paragraph_index,
                footnote_text=footnote_text,
                output_filename=output_filename,
                validate_location=True
            )
        return message
    except Exception as e:
        return f"Failed to add footnote: {str(e)}"


async def customize_footnote_style(filename: str, numbering_format: str = "1, 2, 3", 
                                  start_number: int = 1, font_name: Optional[str] = None,
                                  font_size: Optional[int] = None) -> str:
    """Customize footnote numbering and formatting in a Word document.
    
    Args:
        filename: Path to the Word document
        numbering_format: Format for footnote numbers (e.g., "1, 2, 3", "i, ii, iii", "a, b, c")
        start_number: Number to start footnote numbering from
        font_name: Optional font name for footnotes
        font_size: Optional font size for footnotes (in points)
    """
    filename = ensure_docx_extension(filename)
    
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    
    # Check if file is writeable
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}. Consider creating a copy first."
    
    try:
        async with get_file_lock(filename):
            doc = Document(filename)

            # Create or get footnote style
            footnote_style_name = "Footnote Text"
            footnote_style = None

            try:
                footnote_style = doc.styles[footnote_style_name]
            except KeyError:
                # Create the style if it doesn't exist
                footnote_style = doc.styles.add_style(footnote_style_name, WD_STYLE_TYPE.PARAGRAPH)

            # Apply formatting to footnote style
            if footnote_style:
                if font_name:
                    footnote_style.font.name = font_name
                if font_size:
                    footnote_style.font.size = Pt(font_size)

            # Find all existing footnote references
            footnote_refs = find_footnote_references(doc)

            # Generate format symbols for the specified numbering format
            format_symbols = get_format_symbols(numbering_format, len(footnote_refs) + start_number)

            # Apply custom formatting to footnotes
            count = customize_footnote_formatting(doc, footnote_refs, format_symbols, start_number, footnote_style)

            # Save the document
            doc.save(filename)

        return f"Footnote style and numbering customized in {filename}"
    except Exception as e:
        return f"Failed to customize footnote style: {str(e)}"


async def delete_footnote_from_document(filename: str, footnote_id: Optional[int] = None,
                                       search_text: Optional[str] = None, 
                                       output_filename: Optional[str] = None) -> str:
    """Delete a footnote from a Word document.
    
    You can identify the footnote to delete either by:
    1. footnote_id: The numeric ID of the footnote (1, 2, 3, etc.)
    2. search_text: Text near the footnote reference to find and delete
    
    Args:
        filename: Path to the Word document
        footnote_id: Optional ID of the footnote to delete (1-based)
        search_text: Optional text to search near the footnote reference
        output_filename: Optional output filename (if None, modifies in place)
    """
    filename = ensure_docx_extension(filename)
    
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    
    # Check if file is writeable
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}. Consider creating a copy first."
    
    try:
        async with get_file_lock(filename):
            success, message, details = delete_footnote_robust(
                filename=filename,
                footnote_id=footnote_id,
                search_text=search_text,
                output_filename=output_filename,
                clean_orphans=True
            )
        return message
    except Exception as e:
        return f"Failed to delete footnote: {str(e)}"


# ============================================================================
# Robust tool functions with Dict returns for structured responses
# ============================================================================


async def add_footnote_robust_tool(
    filename: str,
    search_text: Optional[str] = None,
    paragraph_index: Optional[int] = None,
    footnote_text: str = "",
    validate_location: bool = True,
    auto_repair: bool = False
) -> Dict[str, Any]:
    """
    Add a footnote with robust validation and error handling.

    **CANONICAL footnote-add tool — prefer this over `add_footnote_to_document`,
    `add_footnote_enhanced`, `add_footnote_after_text`, and `add_footnote_before_text`.**
    Those older entry points either delegate here or exist only for backward
    compatibility; this one exposes the full feature set (validation, auto-repair,
    structured Dict response) and gives you ID-level control.

    Args:
        filename: Path to the Word document
        search_text: Text to search for (mutually exclusive with paragraph_index)
        paragraph_index: Index of paragraph (mutually exclusive with search_text)
        footnote_text: Content of the footnote
        validate_location: Whether to validate placement restrictions
        auto_repair: Whether to attempt automatic document repair
    
    Returns:
        Dict with success status, message, and optional details
    """
    filename = ensure_docx_extension(filename)
    
    # Check if file is writeable
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return {
            "success": False,
            "message": f"Cannot modify document: {error_message}",
            "details": None
        }
    
    # Convert paragraph_index if provided as string
    if paragraph_index is not None:
        try:
            paragraph_index = int(paragraph_index)
        except (ValueError, TypeError):
            return {
                "success": False,
                "message": "Invalid parameter: paragraph_index must be an integer",
                "details": None
            }
    
    # Call robust implementation
    async with get_file_lock(filename):
        success, message, details = add_footnote_robust(
            filename=filename,
            search_text=search_text,
            paragraph_index=paragraph_index,
            footnote_text=footnote_text,
            validate_location=validate_location,
            auto_repair=auto_repair
        )

    return {
        "success": success,
        "message": message,
        "details": details
    }


async def delete_footnote_robust_tool(
    filename: str,
    footnote_id: Optional[int] = None,
    search_text: Optional[str] = None,
    clean_orphans: bool = True
) -> Dict[str, Any]:
    """
    Delete a footnote with comprehensive cleanup.
    
    Args:
        filename: Path to the Word document
        footnote_id: ID of footnote to delete
        search_text: Text near footnote reference
        clean_orphans: Whether to remove orphaned content
    
    Returns:
        Dict with success status, message, and optional details
    """
    filename = ensure_docx_extension(filename)
    
    # Check if file is writeable
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return {
            "success": False,
            "message": f"Cannot modify document: {error_message}",
            "details": None
        }
    
    # Convert footnote_id if provided as string
    if footnote_id is not None:
        try:
            footnote_id = int(footnote_id)
        except (ValueError, TypeError):
            return {
                "success": False,
                "message": "Invalid parameter: footnote_id must be an integer",
                "details": None
            }
    
    # Call robust implementation
    async with get_file_lock(filename):
        success, message, details = delete_footnote_robust(
            filename=filename,
            footnote_id=footnote_id,
            search_text=search_text,
            clean_orphans=clean_orphans
        )

    return {
        "success": success,
        "message": message,
        "details": details
    }


async def validate_footnotes_tool(filename: str) -> Dict[str, Any]:
    """
    Validate all footnotes in a document.
    
    Provides comprehensive validation report including:
    - ID conflicts
    - Orphaned content
    - Missing styles
    - Invalid locations
    - Coherence issues
    
    Args:
        filename: Path to the Word document
    
    Returns:
        Dict with validation status and detailed report
    """
    filename = ensure_docx_extension(filename)
    
    if not os.path.exists(filename):
        return {
            "valid": False,
            "message": f"Document {filename} does not exist",
            "report": {}
        }
    
    # Call validation
    is_valid, message, report = validate_document_footnotes(filename)
    
    return {
        "valid": is_valid,
        "message": message,
        "report": report
    }


# ============================================================================
# Compatibility wrappers for robust tools (maintain backward compatibility)
# ============================================================================

async def add_footnote_to_document_robust(
    filename: str,
    paragraph_index: int,
    footnote_text: str
) -> str:
    """
    Robust version of add_footnote_to_document.
    Maintains backward compatibility with existing API.
    """
    # Lock acquired inside add_footnote_robust_tool
    result = await add_footnote_robust_tool(
        filename=filename,
        paragraph_index=paragraph_index,
        footnote_text=footnote_text
    )
    return result["message"]


async def add_footnote_after_text_robust(
    filename: str,
    search_text: str,
    footnote_text: str,
    output_filename: Optional[str] = None
) -> str:
    """
    Robust version of add_footnote_after_text.
    Maintains backward compatibility with existing API.
    """
    # Handle output filename by copying first if needed
    working_file = filename
    if output_filename:
        import shutil
        async with get_file_lock(filename):
            shutil.copy2(filename, output_filename)
        working_file = output_filename

    # Lock on working_file acquired inside add_footnote_robust_tool
    result = await add_footnote_robust_tool(
        filename=working_file,
        search_text=search_text,
        footnote_text=footnote_text
    )
    return result["message"]


async def add_footnote_before_text_robust(
    filename: str,
    search_text: str,
    footnote_text: str,
    output_filename: Optional[str] = None
) -> str:
    """
    Robust version of add_footnote_before_text — inserts the footnote
    reference immediately before the matched ``search_text``.
    """
    # Handle output filename
    working_file = filename
    if output_filename:
        import shutil
        async with get_file_lock(filename):
            shutil.copy2(filename, output_filename)
        working_file = output_filename

    # Call the core robust path directly so we can pass position="before".
    # add_footnote_robust_tool() does not expose `position`, so we bypass it.
    async with get_file_lock(working_file):
        success, message, _details = add_footnote_robust(
            filename=working_file,
            search_text=search_text,
            footnote_text=footnote_text,
            position="before",
            validate_location=True,
        )
    return message


async def delete_footnote_from_document_robust(
    filename: str,
    footnote_id: Optional[int] = None,
    search_text: Optional[str] = None,
    output_filename: Optional[str] = None
) -> str:
    """
    Robust version of delete_footnote_from_document.
    Maintains backward compatibility with existing API.
    """
    # Handle output filename
    working_file = filename
    if output_filename:
        import shutil
        async with get_file_lock(filename):
            shutil.copy2(filename, output_filename)
        working_file = output_filename

    # Lock on working_file acquired inside delete_footnote_robust_tool
    result = await delete_footnote_robust_tool(
        filename=working_file,
        footnote_id=footnote_id,
        search_text=search_text
    )
    return result["message"]
