"""
Main entry point for the Word Document MCP Server.
Acts as the central controller for the MCP server that handles Word document operations.
Supports multiple transports: stdio, sse, and streamable-http using standalone FastMCP.
"""

import os
import sys
from dotenv import load_dotenv
from word_document_server.defaults import DEFAULT_AUTHOR, DEFAULT_INITIALS

# Load environment variables from .env file
print("Loading configuration from .env file...", file=sys.stderr)
load_dotenv()
# Set required environment variable for FastMCP 2.8.1+
os.environ.setdefault('FASTMCP_LOG_LEVEL', 'INFO')
from fastmcp import FastMCP
from mcp.types import ToolAnnotations
from word_document_server.tools import (
    document_tools,
    content_tools,
    format_tools,
    protection_tools,
    footnote_tools,
    extended_document_tools,
    comment_tools,
    comment_write_tools,
    hyperlink_tools,
    tracked_changes_tools,
    live_tools,
    live_read_tools,
    live_layout_tools,
    screen_capture_tools,
    layout_tools,
)
from word_document_server.tools.content_tools import replace_paragraph_block_below_header_tool
from word_document_server.tools.content_tools import replace_block_between_manual_anchors_tool
from word_document_server.tools import (
    session_tools,
    window_tools,
    paragraph_format_tools,
    style_tools,
    image_shape_tools,
    table_com_tools,
    vba_tools,
    undo_tools,
    column_layout_tools,
    field_tools,
    mail_merge_tools,
    image_format_tools,
    find_replace_tools,
    navigation_tools,
    list_tools,
    word_diff_tools,
)

def get_transport_config():
    """
    Get transport configuration from environment variables.
    
    Returns:
        dict: Transport configuration with type, host, port, and other settings
    """
    # Default configuration
    config = {
        'transport': 'stdio',  # Default to stdio for backward compatibility
        'host': '0.0.0.0',
        'port': 8000,
        'path': '/mcp',
        'sse_path': '/sse'
    }
    
    # Override with environment variables if provided
    transport = os.getenv('MCP_TRANSPORT', 'stdio').lower()
    print(f"Transport: {transport}", file=sys.stderr)
    # Validate transport type
    valid_transports = ['stdio', 'streamable-http', 'sse']
    if transport not in valid_transports:
        print(f"Warning: Invalid transport '{transport}'. Falling back to 'stdio'.", file=sys.stderr)
        transport = 'stdio'
    
    config['transport'] = transport
    config['host'] = os.getenv('MCP_HOST', config['host'])
    # Use PORT from Render if available, otherwise fall back to MCP_PORT or default
    config['port'] = int(os.getenv('PORT', os.getenv('MCP_PORT', config['port'])))
    config['path'] = os.getenv('MCP_PATH', config['path'])
    config['sse_path'] = os.getenv('MCP_SSE_PATH', config['sse_path'])
    
    return config


def setup_logging(debug_mode):
    """
    Setup logging based on debug mode.
    
    Args:
        debug_mode (bool): Whether to enable debug logging
    """
    import logging
    
    if debug_mode:
        logging.basicConfig(
            level=logging.DEBUG,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        print("Debug logging enabled", file=sys.stderr)
    else:
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )


# Initialize FastMCP server
mcp = FastMCP("Word Document Server")


def register_tools():
    """Register all tools with the MCP server using FastMCP decorators."""
    
    # Document tools (create, copy, info, etc.)
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Create Word Document",
            destructiveHint=True,
        ),
    )
    def create_document(filename: str, title: str = None, author: str = None):
        """Create a new Word document with optional metadata."""
        return document_tools.create_document(filename, title, author)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Copy Word Document",
            destructiveHint=True,
        ),
    )
    def copy_document(source_filename: str, destination_filename: str = None):
        """Create a copy of a Word document."""
        return document_tools.copy_document(source_filename, destination_filename)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get Document Info",
            readOnlyHint=True,
        ),
    )
    def get_document_info(filename: str):
        """Get information about a Word document."""
        return document_tools.get_document_info(filename)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get Document Text",
            readOnlyHint=True,
        ),
    )
    def get_document_text(filename: str, show_revisions: bool = False):
        """Extract all text from a Word document.

        By default returns the effective final text (insertions applied,
        deletions removed).  Set show_revisions=True to get inline redline
        markup where deletions appear as [-deleted-] and insertions as
        {+inserted+}."""
        return document_tools.get_document_text(filename, show_revisions=show_revisions)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get Document Outline",
            readOnlyHint=True,
        ),
    )
    def get_document_outline(filename: str):
        """Get the structure of a Word document."""
        return document_tools.get_document_outline(filename)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="List Available Documents",
            readOnlyHint=True,
        ),
    )
    def list_available_documents(directory: str = "."):
        """List all .docx files in the specified directory."""
        return document_tools.list_available_documents(directory)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get Document XML",
            readOnlyHint=True,
        ),
    )
    def get_document_xml(filename: str):
        """Get the raw XML structure of a Word document."""
        return document_tools.get_document_xml_tool(filename)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Insert Header Near Text",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def insert_header_near_text(filename: str, target_text: str = None, header_title: str = None, position: str = 'after', header_style: str = 'Heading 1', target_paragraph_index: int = None):
        """Insert a header (with specified style) before or after the target paragraph. Specify by text or paragraph index. Args: filename (str), target_text (str, optional), header_title (str), position ('before' or 'after'), header_style (str, default 'Heading 1'), target_paragraph_index (int, optional)."""
        return content_tools.insert_header_near_text_tool(filename, target_text, header_title, position, header_style, target_paragraph_index)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Insert Line Near Text",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def insert_line_or_paragraph_near_text(filename: str, target_text: str = None, line_text: str = None, position: str = 'after', line_style: str = None, target_paragraph_index: int = None):
        """
        Insert a new line or paragraph (with specified or matched style) before or after the target paragraph. Specify by text or paragraph index. Args: filename (str), target_text (str, optional), line_text (str), position ('before' or 'after'), line_style (str, optional), target_paragraph_index (int, optional).
        """
        return content_tools.insert_line_or_paragraph_near_text_tool(filename, target_text, line_text, position, line_style, target_paragraph_index)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Insert List Near Text",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def insert_numbered_list_near_text(filename: str, target_text: str = None, list_items: list[str] = None, position: str = 'after', target_paragraph_index: int = None, bullet_type: str = 'bullet'):
        """Insert a bulleted or numbered list before or after the target paragraph. Specify by text or paragraph index. Args: filename (str), target_text (str, optional), list_items (list of str), position ('before' or 'after'), target_paragraph_index (int, optional), bullet_type ('bullet' for bullets or 'number' for numbered lists, default: 'bullet')."""
        return content_tools.insert_numbered_list_near_text_tool(filename, target_text, list_items, position, target_paragraph_index, bullet_type)
    # Content tools (paragraphs, headings, tables, etc.)
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Paragraph",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_paragraph(filename: str, text: str, style: str = None,
                      font_name: str = None, font_size: int = None,
                      bold: bool = None, italic: bool = None, color: str = None):
        """Add a paragraph to a Word document with optional formatting.

        Args:
            filename: Path to Word document
            text: Paragraph text content
            style: Optional paragraph style name
            font_name: Font family (e.g., 'Helvetica', 'Times New Roman')
            font_size: Font size in points (e.g., 14, 36)
            bold: Make text bold
            italic: Make text italic
            color: Text color as hex RGB (e.g., '000000')
        """
        return content_tools.add_paragraph(filename, text, style, font_name, font_size, bold, italic, color)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Heading",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_heading(filename: str, text: str, level: int = 1,
                    font_name: str = None, font_size: int = None,
                    bold: bool = None, italic: bool = None, border_bottom: bool = False):
        """Add a heading to a Word document with optional formatting.

        Args:
            filename: Path to Word document
            text: Heading text
            level: Heading level (1-9)
            font_name: Font family (e.g., 'Helvetica')
            font_size: Font size in points (e.g., 14)
            bold: Make heading bold
            italic: Make heading italic
            border_bottom: Add bottom border (for section headers)
        """
        return content_tools.add_heading(filename, text, level, font_name, font_size, bold, italic, border_bottom)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Picture",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_picture(filename: str, image_path: str, width: float = None):
        """Add an image to a Word document."""
        return content_tools.add_picture(filename, image_path, width)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Table",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_table(filename: str, rows: int, cols: int, data: list[list[str]] = None):
        """Add a table to a Word document."""
        return content_tools.add_table(filename, rows, cols, data)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Page Break",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_page_break(filename: str):
        """Add a page break to the document."""
        return content_tools.add_page_break(filename)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Delete Paragraph",
            destructiveHint=True,
        ),
    )
    def delete_paragraph(filename: str, paragraph_index: int):
        """Delete a paragraph from a document."""
        return content_tools.delete_paragraph(filename, paragraph_index)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Search and Replace",
            destructiveHint=True,
        ),
    )
    def search_and_replace(filename: str, find_text: str, replace_text: str):
        """Search for text and replace all occurrences."""
        return content_tools.search_and_replace(filename, find_text, replace_text)
    
    # Format tools (styling, text formatting, etc.)
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Create Custom Style",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def create_custom_style(filename: str, style_name: str, bold: bool = None,
                          italic: bool = None, font_size: int = None,
                          font_name: str = None, color: str = None,
                          base_style: str = None):
        """Create a custom style in the document."""
        return format_tools.create_custom_style(
            filename, style_name, bold, italic, font_size, font_name, color, base_style
        )
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Format Text",
            readOnlyHint=False,
            destructiveHint=False,
        ),
        description=format_tools.format_text.__doc__,
    )
    def format_text(filename: str, paragraph_index: int, start_pos: int, end_pos: int,
                   bold: bool = None, italic: bool = None, underline: bool = None,
                   color: str = None, font_size: int = None, font_name: str = None):
        return format_tools.format_text(
            filename, paragraph_index, start_pos, end_pos, bold, italic,
            underline, color, font_size, font_name
        )
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Format Table",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def format_table(filename: str, table_index: int, has_header_row: bool = None,
                    border_style: str = None, shading: list[str] = None):
        """Format a table with borders, shading, and structure."""
        return format_tools.format_table(filename, table_index, has_header_row, border_style, shading)
    
    # New table cell shading tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Set Table Cell Shading",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def set_table_cell_shading(filename: str, table_index: int, row_index: int,
                              col_index: int, fill_color: str, pattern: str = "clear"):
        """Apply shading/filling to a specific table cell."""
        return format_tools.set_table_cell_shading(filename, table_index, row_index, col_index, fill_color, pattern)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Apply Alternating Row Colors",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def apply_table_alternating_rows(filename: str, table_index: int,
                                   color1: str = "FFFFFF", color2: str = "F2F2F2"):
        """Apply alternating row colors to a table for better readability."""
        return format_tools.apply_table_alternating_rows(filename, table_index, color1, color2)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Highlight Table Header",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def highlight_table_header(filename: str, table_index: int,
                             header_color: str = "4472C4", text_color: str = "FFFFFF"):
        """Apply special highlighting to table header row."""
        return format_tools.highlight_table_header(filename, table_index, header_color, text_color)
    
    # Cell merging tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Merge Table Cells",
            readOnlyHint=False,
            destructiveHint=True,
        ),
    )
    def merge_table_cells(filename: str, table_index: int, start_row: int, start_col: int,
                        end_row: int, end_col: int):
        """Merge cells in a rectangular area of a table."""
        return format_tools.merge_table_cells(filename, table_index, start_row, start_col, end_row, end_col)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Merge Cells Horizontally",
            readOnlyHint=False,
            destructiveHint=True,
        ),
    )
    def merge_table_cells_horizontal(filename: str, table_index: int, row_index: int,
                                   start_col: int, end_col: int):
        """Merge cells horizontally in a single row."""
        return format_tools.merge_table_cells_horizontal(filename, table_index, row_index, start_col, end_col)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Merge Cells Vertically",
            readOnlyHint=False,
            destructiveHint=True,
        ),
    )
    def merge_table_cells_vertical(filename: str, table_index: int, col_index: int,
                                 start_row: int, end_row: int):
        """Merge cells vertically in a single column."""
        return format_tools.merge_table_cells_vertical(filename, table_index, col_index, start_row, end_row)
    
    # Cell alignment tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Set Cell Alignment",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def set_table_cell_alignment(filename: str, table_index: int, row_index: int, col_index: int,
                               horizontal: str = "left", vertical: str = "top"):
        """Set text alignment for a specific table cell."""
        return format_tools.set_table_cell_alignment(filename, table_index, row_index, col_index, horizontal, vertical)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Set Table Alignment",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def set_table_alignment_all(filename: str, table_index: int,
                              horizontal: str = "left", vertical: str = "top"):
        """Set text alignment for all cells in a table."""
        return format_tools.set_table_alignment_all(filename, table_index, horizontal, vertical)
    
    # Protection tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Protect Document",
            readOnlyHint=False,
            destructiveHint=True,
        ),
    )
    def protect_document(filename: str, password: str):
        """Add password protection to a Word document."""
        return protection_tools.protect_document(filename, password)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Unprotect Document",
            readOnlyHint=False,
            destructiveHint=True,
        ),
    )
    def unprotect_document(filename: str, password: str):
        """Remove password protection from a Word document."""
        return protection_tools.unprotect_document(filename, password)
    
    # Footnote tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Footnote",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_footnote_to_document(filename: str, paragraph_index: int, footnote_text: str):
        """Add a footnote to a specific paragraph in a Word document."""
        return footnote_tools.add_footnote_to_document(filename, paragraph_index, footnote_text)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Footnote After Text",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_footnote_after_text(filename: str, search_text: str, footnote_text: str,
                               output_filename: str = None):
        """Add a footnote after specific text with proper superscript formatting.
        This enhanced function ensures footnotes display correctly as superscript."""
        return footnote_tools.add_footnote_after_text(filename, search_text, footnote_text, output_filename)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Footnote Before Text",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_footnote_before_text(filename: str, search_text: str, footnote_text: str,
                                output_filename: str = None):
        """Add a footnote before specific text with proper superscript formatting.
        This enhanced function ensures footnotes display correctly as superscript."""
        return footnote_tools.add_footnote_before_text(filename, search_text, footnote_text, output_filename)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Footnote Enhanced",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_footnote_enhanced(filename: str, paragraph_index: int, footnote_text: str,
                             output_filename: str = None):
        """Enhanced footnote addition with guaranteed superscript formatting.
        Adds footnote at the end of a specific paragraph with proper style handling."""
        return footnote_tools.add_footnote_enhanced(filename, paragraph_index, footnote_text, output_filename)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Endnote",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_endnote_to_document(filename: str, paragraph_index: int, endnote_text: str):
        """Add an endnote to a specific paragraph in a Word document."""
        return footnote_tools.add_endnote_to_document(filename, paragraph_index, endnote_text)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Customize Footnote Style",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def customize_footnote_style(filename: str, numbering_format: str = "1, 2, 3",
                                start_number: int = 1, font_name: str = None,
                                font_size: int = None):
        """Customize footnote numbering and formatting in a Word document."""
        return footnote_tools.customize_footnote_style(
            filename, numbering_format, start_number, font_name, font_size
        )
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Delete Footnote",
            destructiveHint=True,
        ),
    )
    def delete_footnote_from_document(filename: str, footnote_id: int = None,
                                     search_text: str = None, output_filename: str = None):
        """Delete a footnote from a Word document.
        Identify the footnote either by ID (1, 2, 3, etc.) or by searching for text near it."""
        return footnote_tools.delete_footnote_from_document(
            filename, footnote_id, search_text, output_filename
        )
    
    # Robust footnote tools - Production-ready with comprehensive validation
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Footnote Robust",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_footnote_robust(filename: str, search_text: str = None,
                           paragraph_index: int = None, footnote_text: str = "",
                           validate_location: bool = True, auto_repair: bool = False):
        """Add footnote with robust validation and Word compliance.
        This is the production-ready version with comprehensive error handling."""
        return footnote_tools.add_footnote_robust_tool(
            filename, search_text, paragraph_index, footnote_text,
            validate_location, auto_repair
        )
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Validate Footnotes",
            readOnlyHint=True,
        ),
    )
    def validate_document_footnotes(filename: str):
        """Validate all footnotes in document for coherence and compliance.
        Returns detailed report on ID conflicts, orphaned content, missing styles, etc."""
        return footnote_tools.validate_footnotes_tool(filename)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Delete Footnote Robust",
            destructiveHint=True,
        ),
    )
    def delete_footnote_robust(filename: str, footnote_id: int = None,
                              search_text: str = None, clean_orphans: bool = True):
        """Delete footnote with comprehensive cleanup and orphan removal.
        Ensures complete removal from document.xml, footnotes.xml, and relationships."""
        return footnote_tools.delete_footnote_robust_tool(
            filename, footnote_id, search_text, clean_orphans
        )
    
    # Extended document tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get Paragraph Text",
            readOnlyHint=True,
        ),
    )
    def get_paragraph_text_from_document(filename: str, paragraph_index: int):
        """Get text from a specific paragraph in a Word document."""
        return extended_document_tools.get_paragraph_text_from_document(filename, paragraph_index)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Find Text",
            readOnlyHint=True,
        ),
    )
    def find_text_in_document(filename: str, text_to_find: str, match_case: bool = True,
                             whole_word: bool = False):
        """Find occurrences of specific text in a Word document."""
        return extended_document_tools.find_text_in_document(
            filename, text_to_find, match_case, whole_word
        )
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get Highlighted Text",
            readOnlyHint=True,
        ),
    )
    def get_highlighted_text(filename: str, color: str = None):
        """Extract all highlighted/colored text from a Word document, including text inside tables."""
        return extended_document_tools.get_highlighted_text_from_document(filename, color)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Convert to PDF",
            destructiveHint=True,
        ),
    )
    def convert_to_pdf(filename: str, output_filename: str = None):
        """Convert a Word document to PDF format."""
        return extended_document_tools.convert_to_pdf(filename, output_filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Replace Block Below Header",
            readOnlyHint=False,
            destructiveHint=True,
        ),
    )
    def replace_paragraph_block_below_header(filename: str, header_text: str, new_paragraphs: list[str], detect_block_end_fn: str = None):
        """Reemplaza el bloque de párrafos debajo de un encabezado, evitando modificar TOC."""
        return replace_paragraph_block_below_header_tool(filename, header_text, new_paragraphs, detect_block_end_fn)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Replace Block Between Anchors",
            readOnlyHint=False,
            destructiveHint=True,
        ),
    )
    def replace_block_between_manual_anchors(filename: str, start_anchor_text: str, new_paragraphs: list[str], end_anchor_text: str = None, match_fn: str = None, new_paragraph_style: str = None):
        """Replace all content between start_anchor_text and end_anchor_text (or next logical header if not provided)."""
        return replace_block_between_manual_anchors_tool(filename, start_anchor_text, new_paragraphs, end_anchor_text, match_fn, new_paragraph_style)

    # Comment tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get All Comments",
            readOnlyHint=True,
        ),
    )
    def get_all_comments(filename: str):
        """Extract all comments from a Word document."""
        return comment_tools.get_all_comments(filename)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get Comments by Author",
            readOnlyHint=True,
        ),
    )
    def get_comments_by_author(filename: str, author: str):
        """Extract comments from a specific author in a Word document."""
        return comment_tools.get_comments_by_author(filename, author)
    
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get Comments for Paragraph",
            readOnlyHint=True,
        ),
    )
    def get_comments_for_paragraph(filename: str, paragraph_index: int):
        """Extract comments for a specific paragraph in a Word document."""
        return comment_tools.get_comments_for_paragraph(filename, paragraph_index)
    # Comment write tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Comment",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_comment(filename: str, target_text: str, comment_text: str,
                    author: str = DEFAULT_AUTHOR, initials: str = DEFAULT_INITIALS):
        """Add a comment to a Word document anchored to specific text.
        The comment will appear in Word's Review panel attached to the target text.

        Args:
            filename: Path to Word document
            target_text: Text in the document to attach the comment to
            comment_text: The comment content
            author: Comment author name (env: MCP_AUTHOR)
            initials: Author initials (env: MCP_AUTHOR_INITIALS)
        """
        return comment_write_tools.add_comment(filename, target_text, comment_text, author, initials)

    # Hyperlink tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Manage Hyperlinks",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def manage_hyperlinks(filename: str, action: str = "add", text: str = "",
                          url: str = "", paragraph_index: int = None):
        """Add or manage hyperlinks in a Word document.
        Finds the specified text and converts it to a clickable hyperlink with blue underline.

        Args:
            filename: Path to Word document
            action: Action to perform ("add" to add a hyperlink)
            text: Text to convert to a hyperlink
            url: URL the hyperlink should point to
            paragraph_index: If specified, only search in this paragraph (0-based)
        """
        return hyperlink_tools.manage_hyperlinks(filename, action, text, url, paragraph_index)

    # New table column width tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Set Column Width",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def set_table_column_width(filename: str, table_index: int, col_index: int,
                              width: float, width_type: str = "points"):
        """Set the width of a specific table column."""
        return format_tools.set_table_column_width(filename, table_index, col_index, width, width_type)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Set Column Widths",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def set_table_column_widths(filename: str, table_index: int, widths: list[float],
                               width_type: str = "points"):
        """Set the widths of multiple table columns."""
        return format_tools.set_table_column_widths(filename, table_index, widths, width_type)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Set Table Width",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def set_table_width(filename: str, table_index: int, width: float,
                       width_type: str = "points"):
        """Set the overall width of a table."""
        return format_tools.set_table_width(filename, table_index, width, width_type)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Auto-Fit Table Columns",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def auto_fit_table_columns(filename: str, table_index: int):
        """Set table columns to auto-fit based on content."""
        return format_tools.auto_fit_table_columns(filename, table_index)

    # New table cell text formatting and padding tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Format Cell Text",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def format_table_cell_text(filename: str, table_index: int, row_index: int, col_index: int,
                               text_content: str = None, bold: bool = None, italic: bool = None,
                               underline: bool = None, color: str = None, font_size: int = None,
                               font_name: str = None):
        """Format text within a specific table cell."""
        return format_tools.format_table_cell_text(filename, table_index, row_index, col_index,
                                                   text_content, bold, italic, underline, color, font_size, font_name)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Set Cell Padding",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def set_table_cell_padding(filename: str, table_index: int, row_index: int, col_index: int,
                               top: float = None, bottom: float = None, left: float = None,
                               right: float = None, unit: str = "points"):
        """Set padding/margins for a specific table cell."""
        return format_tools.set_table_cell_padding(filename, table_index, row_index, col_index,
                                                   top, bottom, left, right, unit)



    # Tracked changes tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Track Replace",
            destructiveHint=True,
        ),
        description=tracked_changes_tools.track_replace.__doc__,
    )
    def track_replace(filename: str, old_text: str, new_text: str, author: str = DEFAULT_AUTHOR):
        return tracked_changes_tools.track_replace(filename, old_text, new_text, author)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Track Insert",
            destructiveHint=True,
        ),
        description=tracked_changes_tools.track_insert.__doc__,
    )
    def track_insert(filename: str, after_text: str, insert_text: str, author: str = DEFAULT_AUTHOR):
        return tracked_changes_tools.track_insert(filename, after_text, insert_text, author)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Track Delete",
            destructiveHint=True,
        ),
        description=tracked_changes_tools.track_delete.__doc__,
    )
    def track_delete(filename: str, text: str, author: str = DEFAULT_AUTHOR):
        return tracked_changes_tools.track_delete(filename, text, author)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="List Tracked Changes",
            readOnlyHint=True,
        ),
    )
    def list_tracked_changes(filename: str):
        """List all tracked changes (insertions and deletions) in a Word document.
        Returns author, date, text, and paragraph context for each change."""
        return tracked_changes_tools.list_tracked_changes(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Accept Tracked Changes",
            destructiveHint=True,
        ),
    )
    def accept_tracked_changes(filename: str, author: str = None, change_ids: list[int] = None):
        """Accept tracked changes: apply insertions (keep text) and remove deletions.
        Optionally filter by author or specific change IDs."""
        return tracked_changes_tools.accept_tracked_changes(filename, author, change_ids)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Reject Tracked Changes",
            destructiveHint=True,
        ),
    )
    def reject_tracked_changes(filename: str, author: str = None, change_ids: list[int] = None):
        """Reject tracked changes: remove insertions and restore deleted text.
        Optionally filter by author or specific change IDs."""
        return tracked_changes_tools.reject_tracked_changes(filename, author, change_ids)

    # --- Live editing tools (Windows + macOS, requires Word running) ---

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Screen Capture",
            readOnlyHint=True,
        ),
    )
    def word_screen_capture(filename: str = None, output_path: str = None):
        """Capture a screenshot of a Word document window.
        Returns the path to the saved PNG image. Requires Word to be running."""
        return screen_capture_tools.word_screen_capture(filename, output_path)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Insert Text",
            destructiveHint=True,
        ),
    )
    def word_live_insert_text(
        filename: str = None,
        text: str = "",
        position: str = "end",
        bookmark: str = None,
        track_changes: bool = False,
    ):
        """Insert text into a Word document that is open in Word.
        Position: 'start', 'end', 'cursor', or character offset. Requires Word running."""
        return live_tools.word_live_insert_text(
            filename, text, position, bookmark, track_changes
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Format Text",
            destructiveHint=True,
        ),
        description=live_tools.word_live_format_text.__doc__,
    )
    def word_live_format_text(
        filename: str = None,
        start: int = None,
        end: int = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        bold: bool = None,
        italic: bool = None,
        underline: bool = None,
        strikethrough: bool = None,
        font_name: str = None,
        font_size: float = None,
        font_color: str = None,
        highlight_color: int = None,
        style_name: str = None,
        paragraph_alignment: str = None,
        page_break_before: bool = None,
        preserve_direct_formatting: bool = False,
        track_changes: bool = False,
    ):
        return live_tools.word_live_format_text(
            filename, start, end, start_paragraph, end_paragraph,
            bold, italic, underline, strikethrough,
            font_name, font_size, font_color, highlight_color,
            style_name, paragraph_alignment, page_break_before,
            preserve_direct_formatting, track_changes,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Replace Text",
            destructiveHint=True,
        ),
        description=live_tools.word_live_replace_text.__doc__,
    )
    def word_live_replace_text(
        filename: str = None,
        find_text: str = "",
        replace_text: str = "",
        match_case: bool = False,
        match_whole_word: bool = False,
        use_wildcards: bool = False,
        replace_all: bool = True,
        track_changes: bool = False,
    ):
        return live_tools.word_live_replace_text(
            filename, find_text, replace_text, match_case,
            match_whole_word, use_wildcards, replace_all, track_changes,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Insert Paragraphs",
            destructiveHint=True,
        ),
        description=live_tools.word_live_insert_paragraphs.__doc__,
    )
    def word_live_insert_paragraphs(
        filename: str = None,
        paragraphs: list = None,
        target_text: str = None,
        target_paragraph_index: int = None,
        position: str = "after",
        style: str = None,
        track_changes: bool = False,
    ):
        return live_tools.word_live_insert_paragraphs(
            filename, paragraphs, target_text, target_paragraph_index,
            position, style, track_changes,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Table",
            destructiveHint=True,
        ),
    )
    def word_live_add_table(
        filename: str = None,
        rows: int = 2,
        cols: int = 2,
        position: str = "end",
        data: list = None,
        style: str = "Table Grid",
        autofit: str = "window",
        track_changes: bool = False,
    ):
        """Add a table to a Word document open in Word.
        Optionally provide data as 2D list. Default style is 'Table Grid' with
        autofit to window width. Set style=None for no style, autofit=None for
        legacy fixed behavior. Requires Word running."""
        return live_tools.word_live_add_table(
            filename, rows, cols, position, data, style, autofit, track_changes
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Format Table",
            destructiveHint=True,
        ),
        description=live_tools.word_live_format_table.__doc__,
    )
    def word_live_format_table(
        filename: str = None,
        table_index: int = -1,
        border_style: str = None,
        cell_bold: list = None,
        cell_alignment: list = None,
        column_widths: list = None,
        table_alignment: str = None,
        cell_shading: list = None,
        autofit: str = None,
    ):
        return live_tools.word_live_format_table(
            filename, table_index, border_style, cell_bold, cell_alignment,
            column_widths, table_alignment, cell_shading, autofit
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Modify Table",
            destructiveHint=True,
        ),
        description=live_tools.word_live_modify_table.__doc__,
    )
    def word_live_modify_table(
        filename: str = None,
        table_index: int = 1,
        operation: str = "get_info",
        row: int = None,
        col: int = None,
        text: str = None,
        before_row: int = None,
        before_col: int = None,
        header: str = None,
        cells: list = None,
        start_row: int = None,
        start_col: int = None,
        end_row: int = None,
        end_col: int = None,
        autofit_mode: str = "content",
        accept_revisions: bool = False,
        track_changes: bool = False,
    ):
        return live_tools.word_live_modify_table(
            filename, table_index, operation, row, col, text,
            before_row, before_col, header, cells,
            start_row, start_col, end_row, end_col,
            autofit_mode, accept_revisions, track_changes,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Delete Text",
            destructiveHint=True,
        ),
    )
    def word_live_delete_text(
        filename: str = None,
        start: int = None,
        end: int = None,
        track_changes: bool = False,
    ):
        """Delete text from a Word document open in Word.
        Specify start/end character positions. Requires Word running."""
        return live_tools.word_live_delete_text(
            filename, start, end, track_changes
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Apply List",
            destructiveHint=True,
        ),
        description=live_tools.word_live_apply_list.__doc__,
    )
    def word_live_apply_list(
        filename: str = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        list_type: str = "bullet",
        level: int = 0,
        remove: bool = False,
        continue_previous: bool = False,
        number_format: dict = None,
        number_style: dict = None,
        start_at: dict = None,
        level_map: dict = None,
        track_changes: bool = False,
    ):
        return live_tools.word_live_apply_list(
            filename, start_paragraph, end_paragraph, list_type,
            level, remove, continue_previous, number_format,
            number_style, start_at, level_map, track_changes,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Setup Heading Numbering",
            destructiveHint=True,
        ),
        description=live_tools.word_live_setup_heading_numbering.__doc__,
    )
    def word_live_setup_heading_numbering(
        filename: str = None,
        h1_paragraphs: list = None,
        h2_paragraphs: list = None,
        strip_manual_numbers: bool = True,
        h1_number_format: str = None,
        h2_number_format: str = None,
        font_name: str = None,
        h1_size: float = None,
        h2_size: float = None,
        bold: bool = None,
        alignment: str = None,
        font_color: str = None,
        h1_space_before: float = None,
        h1_space_after: float = None,
        h2_space_before: float = None,
        h2_space_after: float = None,
        line_spacing: float = None,
    ):
        return live_tools.word_live_setup_heading_numbering(
            filename, h1_paragraphs, h2_paragraphs, strip_manual_numbers,
            h1_number_format, h2_number_format,
            font_name, h1_size, h2_size, bold, alignment, font_color,
            h1_space_before, h1_space_after, h2_space_before, h2_space_after,
            line_spacing,
        )

    # --- Live read tools (Windows only, requires Word running) ---

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Get Text",
            readOnlyHint=True,
        ),
    )
    def word_live_get_text(filename: str = None):
        """Get all text from a Word document open in Word, paragraph by paragraph. For large documents (200+ paragraphs), automatically returns only the first 3 pages — use word_live_get_page_text to read specific pages."""
        return live_read_tools.word_live_get_text(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Take Snapshot",
            readOnlyHint=True,
        ),
    )
    def word_live_take_snapshot(filename: str = None):
        """Store a snapshot of the current document text for later diffing without returning the full text. Use word_live_get_diff afterwards to see what changed."""
        return live_read_tools.word_live_take_snapshot(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Get Diff",
            readOnlyHint=True,
        ),
    )
    def word_live_get_diff(filename: str = None):
        """Return only paragraphs that changed since the last snapshot. Compares current document against snapshot from word_live_take_snapshot. Returns added, modified, deleted paragraphs. Automatically updates snapshot after diffing."""
        return live_read_tools.word_live_get_diff(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Snapshot Status",
            readOnlyHint=True,
        ),
    )
    def word_live_snapshot_status(filename: str = None):
        """Check whether a snapshot exists for the document and how old it is. Returns has_snapshot, age_seconds, and paragraph_count."""
        return live_read_tools.word_live_snapshot_status(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Get Paragraph Format",
            readOnlyHint=True,
        ),
        description=live_read_tools.word_live_get_paragraph_format.__doc__,
    )
    def word_live_get_paragraph_format(
        filename: str = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        include_runs: bool = False,
    ):
        return live_read_tools.word_live_get_paragraph_format(
            filename, start_paragraph, end_paragraph, include_runs,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Get Info",
            readOnlyHint=True,
        ),
    )
    def word_live_get_info(filename: str = None):
        """Get document info (pages, words, sections, etc.) from a Word document open in Word. Requires Word running."""
        return live_read_tools.word_live_get_info(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Set Core Properties",
            destructiveHint=True,
        ),
        description=live_read_tools.word_live_set_core_properties.__doc__,
    )
    def word_live_set_core_properties(
        filename: str = None,
        title: str = None,
        subject: str = None,
        author: str = None,
        keywords: str = None,
        comments: str = None,
        category: str = None,
        manager: str = None,
        company: str = None,
        last_author: str = None,
    ):
        return live_read_tools.word_live_set_core_properties(
            filename=filename,
            title=title,
            subject=subject,
            author=author,
            keywords=keywords,
            comments=comments,
            category=category,
            manager=manager,
            company=company,
            last_author=last_author,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live List Open",
            readOnlyHint=True,
        ),
    )
    def word_live_list_open():
        """List all documents currently open in Word with name, path, pages, and saved status."""
        return live_read_tools.word_live_list_open()

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Find Text",
            readOnlyHint=True,
        ),
    )
    def word_live_find_text(
        filename: str = None,
        search_text: str = "",
        match_case: bool = False,
        whole_word: bool = False,
        use_wildcards: bool = False,
        context_chars: int = 60,
        max_results: int = 50,
    ):
        """Find text in a Word document open in Word. Returns positions and context.
        With use_wildcards=True, supports ^m (page break), ^t (tab), ^p (paragraph mark) and Word wildcards.
        context_chars controls how many characters of surrounding context to return (default 60). Requires Word running."""
        return live_read_tools.word_live_find_text(
            filename, search_text, match_case, whole_word,
            use_wildcards, context_chars, max_results,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Get Comments",
            readOnlyHint=True,
        ),
    )
    def word_live_get_comments(filename: str = None):
        """Get all comments from a Word document open in Word. Requires Word running."""
        return live_read_tools.word_live_get_comments(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Comment",
            destructiveHint=True,
        ),
    )
    def word_live_add_comment(
        filename: str = None,
        start: int = None,
        end: int = None,
        paragraph_index: int = None,
        text: str = "",
        author: str = DEFAULT_AUTHOR,
    ):
        """Add a comment to a Word document open in Word.
        Specify start/end character positions or paragraph_index (1-indexed). Requires Word running."""
        return live_read_tools.word_live_add_comment(
            filename, start, end, paragraph_index, text, author
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Reply to Comment",
            destructiveHint=True,
        ),
    )
    def word_live_reply_to_comment(
        filename: str = None,
        comment_index: int = None,
        text: str = "",
        author: str = DEFAULT_AUTHOR,
    ):
        """Reply to an existing comment in a Word document open in Word.
        Adds a threaded reply. Use word_live_get_comments to find the comment_index.
        Requires Word 2016+ running."""
        return live_read_tools.word_live_reply_to_comment(
            filename, comment_index, text, author
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Resolve Comment",
            destructiveHint=True,
        ),
    )
    def word_live_resolve_comment(
        filename: str = None,
        comment_index: int = None,
        resolve: bool = True,
    ):
        """Resolve or unresolve a comment in a Word document open in Word.
        Sets the comment's Done property. Use word_live_get_comments to find comment_index.
        Requires Word 2016+ running."""
        return live_read_tools.word_live_resolve_comment(
            filename, comment_index, resolve
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Delete Comment",
            destructiveHint=True,
        ),
    )
    def word_live_delete_comment(
        filename: str = None,
        comment_index: int = None,
    ):
        """Delete a comment from a Word document open in Word.
        Permanently removes the comment. Use word_live_get_comments to find comment_index.
        Requires Word running."""
        return live_read_tools.word_live_delete_comment(
            filename, comment_index
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live List Revisions",
            readOnlyHint=True,
        ),
    )
    def word_live_list_revisions(filename: str = None):
        """List all tracked changes (revisions) in a Word document open in Word. Requires Word running."""
        return live_read_tools.word_live_list_revisions(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Accept Revisions",
            destructiveHint=True,
        ),
    )
    def word_live_accept_revisions(
        filename: str = None,
        author: str = None,
        revision_ids: list[int] = None,
    ):
        """Accept tracked changes in a Word document open in Word.
        Filter by author or specific revision IDs. Requires Word running."""
        return live_read_tools.word_live_accept_revisions(
            filename, author, revision_ids
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Reject Revisions",
            destructiveHint=True,
        ),
    )
    def word_live_reject_revisions(
        filename: str = None,
        author: str = None,
        revision_ids: list[int] = None,
    ):
        """Reject tracked changes in a Word document open in Word.
        Filter by author or specific revision IDs. Requires Word running."""
        return live_read_tools.word_live_reject_revisions(
            filename, author, revision_ids
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Get Page Text",
            readOnlyHint=True,
        ),
        description=live_read_tools.word_live_get_page_text.__doc__,
    )
    def word_live_get_page_text(
        filename: str = None,
        page: int = 1,
        end_page: int = None,
    ):
        return live_read_tools.word_live_get_page_text(
            filename, page, end_page,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Get Undo History",
            readOnlyHint=True,
        ),
    )
    def word_live_get_undo_history(filename: str = None):
        """Get the undo stack from a Word document open in Word.
        Shows MCP tool operations as named entries. Requires Word running."""
        return live_read_tools.word_live_get_undo_history(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Undo",
            destructiveHint=True,
        ),
    )
    def word_live_undo(
        filename: str = None,
        times: int = 1,
    ):
        """Undo the last N operations in a Word document open in Word.
        Each MCP tool call is one undo entry. Requires Word running."""
        return live_tools.word_live_undo(filename, times)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Save",
            destructiveHint=True,
        ),
    )
    def word_live_save(
        filename: str = None,
        save_as: str = None,
    ):
        """Save a Word document open in Word.
        Optionally save to a new path with save_as. Requires Word running."""
        return live_tools.word_live_save(filename, save_as)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Toggle Track Changes",
            destructiveHint=True,
        ),
    )
    def word_live_toggle_track_changes(
        filename: str = None,
        enable: bool = None,
    ):
        """Toggle or set Track Changes mode on a Word document.
        If enable is omitted, toggles current state. Requires Word running."""
        return live_tools.word_live_toggle_track_changes(filename, enable)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Insert Image",
            destructiveHint=True,
        ),
        description=live_tools.word_live_insert_image.__doc__,
    )
    def word_live_insert_image(
        filename: str = None,
        image_path: str = "",
        paragraph_index: int = None,
        position: str = "end",
        width_inches: float = None,
        height_inches: float = None,
        width_pt: float = None,
        height_pt: float = None,
        alignment: str = None,
        wrapping: str = None,
        border_style: str = None,
        border_width_pt: float = None,
        border_color: str = None,
        link_to_file: bool = False,
    ):
        return live_tools.word_live_insert_image(
            filename, image_path, paragraph_index, position,
            width_inches, height_inches, width_pt, height_pt,
            alignment, wrapping, border_style, border_width_pt,
            border_color, link_to_file
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Insert Cross Reference",
            destructiveHint=True,
        ),
    )
    def word_live_insert_cross_reference(
        filename: str = None,
        ref_type: str = "heading",
        ref_item: int = 1,
        ref_kind: str = "text",
        insert_position: str = "end",
        paragraph_index: int = None,
        insert_as_hyperlink: bool = True,
    ):
        """Insert a cross-reference to a heading, bookmark, figure, table, etc.
        First use word_live_list_cross_reference_items to discover available targets.
        ref_type: heading, bookmark, figure, table, equation, footnote, endnote.
        ref_kind: text, number, number_no_context, page, above_below.
        Requires Word running."""
        return live_tools.word_live_insert_cross_reference(
            filename, ref_type, ref_item, ref_kind,
            insert_position, paragraph_index, insert_as_hyperlink
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live List Cross Reference Items",
            readOnlyHint=True,
        ),
    )
    def word_live_list_cross_reference_items(
        filename: str = None,
        ref_type: str = "heading",
    ):
        """List available cross-reference targets in a Word document.
        Returns items that can be referenced with word_live_insert_cross_reference.
        ref_type: heading, bookmark, figure, table, equation, footnote, endnote.
        Requires Word running."""
        return live_tools.word_live_list_cross_reference_items(filename, ref_type)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Insert Equation",
            destructiveHint=True,
        ),
    )
    def word_live_insert_equation(
        filename: str = None,
        equation: str = "",
        paragraph_index: int = None,
        position: str = "end",
        display_mode: bool = False,
    ):
        """Insert a mathematical equation into a Word document using UnicodeMath syntax.
        Examples: "x^2 + y^2 = z^2", "(a+b)/(c+d)" (fraction), "\\sqrt(x^2+y^2)" (root),
        "\\alpha + \\beta" (Greek), "\\int_0^\\infty e^(-x^2) dx" (integral),
        "\\sum_(i=1)^n i^2" (summation), "\\matrix(a&b@c&d)" (matrix).
        display_mode=True centers the equation on its own line.
        Requires Word running."""
        return live_tools.word_live_insert_equation(
            filename, equation, paragraph_index, position, display_mode
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Diagnose Layout",
            readOnlyHint=True,
        ),
        description=live_read_tools.word_live_diagnose_layout.__doc__,
    )
    def word_live_diagnose_layout(filename: str = None):
        return live_read_tools.word_live_diagnose_layout(filename)

    # --- Live layout tools (Windows only, requires Word running) ---

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Set Page Layout",
            destructiveHint=True,
        ),
    )
    def word_live_set_page_layout(
        filename: str = None,
        section_index: int = 1,
        orientation: str = None,
        page_width_inches: float = None,
        page_height_inches: float = None,
        margin_top_inches: float = None,
        margin_bottom_inches: float = None,
        margin_left_inches: float = None,
        margin_right_inches: float = None,
    ):
        """Set page layout (orientation, size, margins) for a section in a Word document open in Word. Requires Word running."""
        return live_layout_tools.word_live_set_page_layout(
            filename, section_index, orientation,
            page_width_inches, page_height_inches,
            margin_top_inches, margin_bottom_inches,
            margin_left_inches, margin_right_inches,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Header/Footer",
            destructiveHint=True,
        ),
    )
    def word_live_add_header_footer(
        filename: str = None,
        section_index: int = 1,
        header_text: str = None,
        footer_text: str = None,
        header_alignment: str = "center",
        footer_alignment: str = "center",
    ):
        """Add header and/or footer to a section in a Word document open in Word. Requires Word running."""
        return live_layout_tools.word_live_add_header_footer(
            filename, section_index, header_text, footer_text,
            header_alignment, footer_alignment,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Page Numbers",
            destructiveHint=True,
        ),
    )
    def word_live_add_page_numbers(
        filename: str = None,
        section_index: int = 1,
        position: str = "footer",
        alignment: str = "center",
        prefix: str = "",
        suffix: str = "",
        include_total: bool = False,
    ):
        """Add page numbers to header or footer in a Word document open in Word. Requires Word running."""
        return live_layout_tools.word_live_add_page_numbers(
            filename, section_index, position, alignment,
            prefix, suffix, include_total,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Section Break",
            destructiveHint=True,
        ),
    )
    def word_live_add_section_break(
        filename: str = None,
        break_type: str = "new_page",
    ):
        """Add a section break (new_page, continuous, even_page, odd_page) to a Word document open in Word. Requires Word running."""
        return live_layout_tools.word_live_add_section_break(
            filename, break_type,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Set Paragraph Spacing",
            destructiveHint=True,
        ),
    )
    def word_live_set_paragraph_spacing(
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
    ):
        """Set paragraph spacing and layout properties in a Word document open in Word. Paragraphs are 1-indexed. Requires Word running."""
        return live_layout_tools.word_live_set_paragraph_spacing(
            filename, paragraph_index, start_paragraph, end_paragraph,
            space_before_pt, space_after_pt, line_spacing, line_spacing_rule,
            keep_with_next, keep_together, alignment,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Bookmark",
            destructiveHint=True,
        ),
    )
    def word_live_add_bookmark(
        filename: str = None,
        paragraph_index: int = 1,
        bookmark_name: str = "",
    ):
        """Add a named bookmark at a paragraph in a Word document open in Word.
        Paragraph is 1-indexed. Requires Word running."""
        return live_layout_tools.word_live_add_bookmark(
            filename, paragraph_index, bookmark_name,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Watermark",
            destructiveHint=True,
        ),
    )
    def word_live_add_watermark(
        filename: str = None,
        text: str = "TASLAK",
        font_size: int = 72,
        font_color: str = "C0C0C0",
        rotation: int = -45,
        section_index: int = 1,
    ):
        """Add a diagonal text watermark to a Word document open in Word. Requires Word running."""
        return live_layout_tools.word_live_add_watermark(
            filename, text, font_size, font_color, rotation, section_index,
        )

    # --- Layout, header/footer, spacing, bookmark, watermark tools ---

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Set Page Layout",
            destructiveHint=True,
        ),
    )
    def set_page_layout(
        filename: str,
        section_index: int = 0,
        orientation: str = None,
        page_width_inches: float = None,
        page_height_inches: float = None,
        margin_top_inches: float = None,
        margin_bottom_inches: float = None,
        margin_left_inches: float = None,
        margin_right_inches: float = None,
    ):
        """Set page layout (orientation, size, margins) for a document section."""
        return layout_tools.set_page_layout(
            filename, section_index, orientation,
            page_width_inches, page_height_inches,
            margin_top_inches, margin_bottom_inches,
            margin_left_inches, margin_right_inches,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Header/Footer",
            destructiveHint=True,
        ),
    )
    def add_header_footer(
        filename: str,
        section_index: int = 0,
        header_text: str = None,
        footer_text: str = None,
        header_alignment: str = "center",
        footer_alignment: str = "center",
    ):
        """Add header and/or footer text to a document section."""
        return layout_tools.add_header_footer(
            filename, section_index, header_text, footer_text,
            header_alignment, footer_alignment,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Page Numbers",
            destructiveHint=True,
        ),
    )
    def add_page_numbers(
        filename: str,
        section_index: int = 0,
        position: str = "footer",
        alignment: str = "center",
        prefix: str = "",
        suffix: str = "",
        include_total: bool = False,
    ):
        """Add page numbers to header or footer using PAGE/NUMPAGES fields."""
        return layout_tools.add_page_numbers(
            filename, section_index, position, alignment,
            prefix, suffix, include_total,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Section Break",
            destructiveHint=True,
        ),
    )
    def add_section_break(filename: str, break_type: str = "new_page"):
        """Add a section break (new_page, continuous, even_page, odd_page)."""
        return layout_tools.add_section_break(filename, break_type)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Set Paragraph Spacing",
            destructiveHint=True,
        ),
    )
    def set_paragraph_spacing(
        filename: str,
        paragraph_index: int = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        space_before_pt: float = None,
        space_after_pt: float = None,
        line_spacing: float = None,
        line_spacing_rule: str = None,
    ):
        """Set paragraph spacing (before/after/line) for one or a range of paragraphs.
        line_spacing_rule: single, 1.5_lines, double, exactly, at_least, multiple."""
        return layout_tools.set_paragraph_spacing(
            filename, paragraph_index, start_paragraph, end_paragraph,
            space_before_pt, space_after_pt, line_spacing, line_spacing_rule,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Bookmark",
            destructiveHint=True,
        ),
    )
    def add_bookmark(filename: str, paragraph_index: int, bookmark_name: str):
        """Add a named bookmark at a paragraph for cross-referencing."""
        return layout_tools.add_bookmark(filename, paragraph_index, bookmark_name)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Watermark",
            destructiveHint=True,
        ),
    )
    def add_watermark(
        filename: str,
        text: str = "TASLAK",
        font_size: int = 72,
        font_color: str = "C0C0C0",
        rotation: int = -45,
        section_index: int = 0,
    ):
        """Add a diagonal text watermark (e.g. TASLAK, GİZLİ, DRAFT) to a document."""
        return layout_tools.add_watermark(
            filename, text, font_size, font_color, rotation, section_index,
        )

    # --- Previously unregistered existing tools ---

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Table of Contents",
            destructiveHint=True,
        ),
    )
    def add_table_of_contents(filename: str, title: str = "Table of Contents", max_level: int = 3):
        """Add a table of contents based on heading styles."""
        return content_tools.add_table_of_contents(filename, title, max_level)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Merge Documents",
            destructiveHint=True,
        ),
    )
    def merge_documents(target_filename: str, source_filenames: list[str], add_page_breaks: bool = True):
        """Merge multiple Word documents into a single target document."""
        return document_tools.merge_documents(target_filename, source_filenames, add_page_breaks)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Restricted Editing",
            destructiveHint=True,
        ),
    )
    def add_restricted_editing(filename: str, password: str, editable_sections: list[str]):
        """Add restricted editing to a document, allowing editing only in specified sections."""
        return protection_tools.add_restricted_editing(filename, password, editable_sections)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Digital Signature",
            destructiveHint=True,
        ),
    )
    def add_digital_signature(filename: str, signer_name: str, reason: str = None):
        """Add a digital signature to a Word document."""
        return protection_tools.add_digital_signature(filename, signer_name, reason)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Verify Document",
            readOnlyHint=True,
        ),
    )
    def verify_document(filename: str, password: str = None):
        """Verify document protection and/or digital signature."""
        return protection_tools.verify_document(filename, password)

    # ── Session management ────────────────────────────────────────────────────

    @mcp.tool(annotations=ToolAnnotations(title="Open Document in Word", destructiveHint=False))
    async def word_open_document(path: str, visible: bool = True, read_only: bool = False, password: str = None):
        """Open a .docx/.doc file in the live Word application. Launches Word automatically if not running."""
        return await session_tools.word_open_document(path, visible, read_only, password)

    @mcp.tool(annotations=ToolAnnotations(title="New Document in Word", destructiveHint=False))
    async def word_new_document(template: str = None, visible: bool = True, title: str = None):
        """Create a new blank Word document (optionally from a template .dotx file)."""
        return await session_tools.word_new_document(template, visible, title)

    @mcp.tool(annotations=ToolAnnotations(title="Close Document", destructiveHint=True))
    async def word_close_document(filename: str = None, save: bool = True, save_path: str = None):
        """Close an open Word document. save=True saves in place; provide save_path for Save As."""
        return await session_tools.word_close_document(filename, save, save_path)

    @mcp.tool(annotations=ToolAnnotations(title="Save Document", destructiveHint=True))
    async def word_save_document(filename: str = None, save_path: str = None, format: str = "docx"):
        """Save an open Word document. format: docx | doc | pdf | txt | rtf."""
        return await session_tools.word_save_document(filename, save_path, format)

    @mcp.tool(annotations=ToolAnnotations(title="List Open Documents", readOnlyHint=True))
    async def word_list_sessions():
        """List all documents currently open in Word with their names, paths, and modification state."""
        return await session_tools.word_list_sessions()

    @mcp.tool(annotations=ToolAnnotations(title="Activate Document", destructiveHint=False))
    async def word_activate_document(filename: str):
        """Bring a specific open document to the foreground in Word."""
        return await session_tools.word_activate_document(filename)

    # ── Window management ─────────────────────────────────────────────────────

    @mcp.tool(annotations=ToolAnnotations(title="Show Word Window", destructiveHint=False))
    async def word_show(filename: str = None):
        """Make the Word application window visible and bring it to the foreground."""
        return await window_tools.word_show(filename)

    @mcp.tool(annotations=ToolAnnotations(title="Hide Word Window", destructiveHint=False))
    async def word_hide():
        """Hide the Word application window (Word keeps running in background)."""
        return await window_tools.word_hide()

    @mcp.tool(annotations=ToolAnnotations(title="Get Word Window Info", readOnlyHint=True))
    async def word_get_window_info():
        """Get the current Word window state, position, and size."""
        return await window_tools.word_get_window_info()

    @mcp.tool(annotations=ToolAnnotations(title="Set Word Window State", destructiveHint=False))
    async def word_set_window_state(state: str):
        """Set Word window state: 'normal', 'minimized', or 'maximized'."""
        return await window_tools.word_set_window_state(state)

    @mcp.tool(annotations=ToolAnnotations(title="Set Word Window Position", destructiveHint=False))
    async def word_set_window_position(left: float, top: float, width: float = None, height: float = None):
        """Set Word window position and optionally size (in points)."""
        return await window_tools.word_set_window_position(left, top, width, height)

    @mcp.tool(annotations=ToolAnnotations(title="Set Word Status Bar", destructiveHint=False))
    async def word_set_status_bar(text: str):
        """Display a message in the Word status bar."""
        return await window_tools.word_set_status_bar(text)

    # ── Paragraph formatting ──────────────────────────────────────────────────

    @mcp.tool(annotations=ToolAnnotations(title="Set Paragraph Indent", destructiveHint=True))
    async def word_live_set_paragraph_indent(
        filename: str = None,
        paragraph_index: int = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        left_indent_cm: float = None,
        right_indent_cm: float = None,
        first_line_indent_cm: float = None,
        hanging_indent_cm: float = None,
    ):
        """Set paragraph indentation (left, right, first-line, hanging) in cm."""
        return await paragraph_format_tools.word_live_set_paragraph_indent(
            filename, paragraph_index, start_paragraph, end_paragraph,
            left_indent_cm, right_indent_cm, first_line_indent_cm, hanging_indent_cm,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Set Tab Stops", destructiveHint=True))
    async def word_live_set_tab_stops(
        filename: str = None,
        paragraph_index: int = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        tabs: list = None,
        clear_existing: bool = False,
    ):
        """Set custom tab stops on paragraphs. tabs = [{position_cm, alignment, leader}, ...]."""
        return await paragraph_format_tools.word_live_set_tab_stops(
            filename, paragraph_index, start_paragraph, end_paragraph, tabs, clear_existing,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Set Paragraph Border", destructiveHint=True))
    async def word_live_set_paragraph_border(
        filename: str = None,
        paragraph_index: int = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        sides: list = None,
        line_style: str = "single",
        line_width_pt: float = 0.75,
        color: str = "#000000",
    ):
        """Add borders to paragraphs. sides: list of 'top'|'bottom'|'left'|'right'|'all'."""
        return await paragraph_format_tools.word_live_set_paragraph_border(
            filename, paragraph_index, start_paragraph, end_paragraph,
            sides, line_style, line_width_pt, color,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Clear Paragraph Border", destructiveHint=True))
    async def word_live_clear_paragraph_border(
        filename: str = None,
        paragraph_index: int = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
    ):
        """Remove all borders from paragraphs."""
        return await paragraph_format_tools.word_live_clear_paragraph_border(
            filename, paragraph_index, start_paragraph, end_paragraph,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Set Paragraph Shading", destructiveHint=True))
    async def word_live_set_paragraph_shading(
        filename: str = None,
        paragraph_index: int = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        background_color: str = None,
        foreground_color: str = None,
        pattern: str = "solid",
    ):
        """Set paragraph background shading/fill color. pattern: solid|clear|5_percent|..."""
        return await paragraph_format_tools.word_live_set_paragraph_shading(
            filename, paragraph_index, start_paragraph, end_paragraph,
            background_color, foreground_color, pattern,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Set Character Spacing", destructiveHint=True))
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
    ):
        """Set letter spacing, horizontal scaling, vertical position, and kerning."""
        return await paragraph_format_tools.word_live_set_character_spacing(
            filename, start, end, start_paragraph, end_paragraph,
            spacing_pt, scale_percent, position_pt, kerning_pt,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Set Text Effects", destructiveHint=True))
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
    ):
        """Apply text effects: superscript, subscript, small caps, all caps, shadow, outline, etc."""
        return await paragraph_format_tools.word_live_set_text_effects(
            filename, start, end, start_paragraph, end_paragraph,
            superscript, subscript, small_caps, all_caps, hidden,
            double_strikethrough, outline, shadow, emboss, engrave,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Clear Formatting", destructiveHint=True))
    async def word_live_clear_formatting(
        filename: str = None,
        start: int = None,
        end: int = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        clear_char_formatting: bool = True,
        clear_para_formatting: bool = True,
    ):
        """Clear all direct formatting from a range, resetting to style defaults."""
        return await paragraph_format_tools.word_live_clear_formatting(
            filename, start, end, start_paragraph, end_paragraph,
            clear_char_formatting, clear_para_formatting,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Get Paragraph Formatting", readOnlyHint=True))
    async def word_live_get_paragraph_formatting(filename: str = None, paragraph_index: int = 1):
        """Inspect all formatting of a paragraph: font, alignment, indent, spacing, borders, style."""
        return await paragraph_format_tools.word_live_get_paragraph_formatting(filename, paragraph_index)

    # ── Style management ──────────────────────────────────────────────────────

    @mcp.tool(annotations=ToolAnnotations(title="List Styles", readOnlyHint=True))
    async def word_live_list_styles(filename: str = None, style_type: str = "all", include_builtin: bool = True, include_custom: bool = True):
        """List available styles. style_type: all | paragraph | character | table | list."""
        return await style_tools.word_live_list_styles(filename, style_type, include_builtin, include_custom)

    @mcp.tool(annotations=ToolAnnotations(title="Apply Paragraph Style", destructiveHint=True))
    async def word_live_apply_paragraph_style(
        filename: str = None,
        paragraph_index: int = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        style_name: str = "Normal",
    ):
        """Apply a paragraph style (e.g. 'Heading 1', 'Normal', 'List Paragraph') to paragraphs."""
        return await style_tools.word_live_apply_paragraph_style(
            filename, paragraph_index, start_paragraph, end_paragraph, style_name,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Apply Character Style", destructiveHint=True))
    async def word_live_apply_character_style(
        filename: str = None,
        start: int = None,
        end: int = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        style_name: str = "Default Paragraph Font",
    ):
        """Apply a character style (e.g. 'Strong', 'Emphasis') to a text range."""
        return await style_tools.word_live_apply_character_style(
            filename, start, end, start_paragraph, end_paragraph, style_name,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Create Style", destructiveHint=True))
    async def word_live_create_style(
        filename: str = None,
        style_name: str = "",
        style_type: str = "paragraph",
        base_style: str = "Normal",
        font_name: str = None,
        font_size: float = None,
        bold: bool = None,
        italic: bool = None,
        font_color: str = None,
        alignment: str = None,
        space_before_pt: float = None,
        space_after_pt: float = None,
        left_indent_cm: float = None,
    ):
        """Create a new custom paragraph or character style with optional font/spacing properties."""
        return await style_tools.word_live_create_style(
            filename, style_name, style_type, base_style,
            font_name, font_size, bold, italic, font_color,
            alignment, space_before_pt, space_after_pt, left_indent_cm,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Delete Style", destructiveHint=True))
    async def word_live_delete_style(filename: str = None, style_name: str = ""):
        """Delete a custom style from the document. Built-in styles cannot be deleted."""
        return await style_tools.word_live_delete_style(filename, style_name)

    @mcp.tool(annotations=ToolAnnotations(title="Rename Style", destructiveHint=True))
    async def word_live_rename_style(filename: str = None, old_name: str = "", new_name: str = ""):
        """Rename a custom style in the document."""
        return await style_tools.word_live_rename_style(filename, old_name, new_name)

    # ── Image & shape tools ───────────────────────────────────────────────────

    @mcp.tool(annotations=ToolAnnotations(title="List Images", readOnlyHint=True))
    async def word_live_list_images(filename: str = None):
        """List all inline images with index, size, alt text, and paragraph position."""
        return await image_shape_tools.word_live_list_images(filename)

    @mcp.tool(annotations=ToolAnnotations(title="Delete Image", destructiveHint=True))
    async def word_live_delete_image(filename: str = None, image_index: int = 1):
        """Delete an inline image by its 1-based index (use word_live_list_images to find it)."""
        return await image_shape_tools.word_live_delete_image(filename, image_index)

    @mcp.tool(annotations=ToolAnnotations(title="Resize Image", destructiveHint=True))
    async def word_live_resize_image(filename: str = None, image_index: int = 1, width_cm: float = None, height_cm: float = None, lock_aspect_ratio: bool = True):
        """Resize an inline image. Provide width_cm or height_cm (or both). Aspect ratio locked by default."""
        return await image_shape_tools.word_live_resize_image(filename, image_index, width_cm, height_cm, lock_aspect_ratio)

    @mcp.tool(annotations=ToolAnnotations(title="List Shapes", readOnlyHint=True))
    async def word_live_list_shapes(filename: str = None):
        """List all floating shapes (text boxes, drawings) with index, name, position, and text."""
        return await image_shape_tools.word_live_list_shapes(filename)

    @mcp.tool(annotations=ToolAnnotations(title="Delete Shape", destructiveHint=True))
    async def word_live_delete_shape(filename: str = None, shape_index: int = None, shape_name: str = None):
        """Delete a floating shape by index or name (use word_live_list_shapes to find it)."""
        return await image_shape_tools.word_live_delete_shape(filename, shape_index, shape_name)

    @mcp.tool(annotations=ToolAnnotations(title="Add Text Box", destructiveHint=True))
    async def word_live_add_text_box(
        filename: str = None,
        text: str = "",
        left_cm: float = 2.0,
        top_cm: float = 2.0,
        width_cm: float = 8.0,
        height_cm: float = 3.0,
        font_name: str = None,
        font_size: float = None,
        bold: bool = None,
        italic: bool = None,
        font_color: str = None,
        border: bool = True,
        fill_color: str = None,
    ):
        """Add a floating text box at a specified position with optional font and fill settings."""
        return await image_shape_tools.word_live_add_text_box(
            filename, text, left_cm, top_cm, width_cm, height_cm,
            font_name, font_size, bold, italic, font_color, border, fill_color,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Move/Resize Shape", destructiveHint=True))
    async def word_live_move_shape(filename: str = None, shape_index: int = None, shape_name: str = None, left_cm: float = None, top_cm: float = None, width_cm: float = None, height_cm: float = None):
        """Move or resize a floating shape by index or name."""
        return await image_shape_tools.word_live_move_shape(filename, shape_index, shape_name, left_cm, top_cm, width_cm, height_cm)

    # ── Table COM tools ───────────────────────────────────────────────────────

    @mcp.tool(annotations=ToolAnnotations(title="Get Table Data", readOnlyHint=True))
    async def word_live_get_table_data(filename: str = None, table_index: int = 1, include_formatting: bool = False):
        """Read all cell values from a table as a 2D array."""
        return await table_com_tools.word_live_get_table_data(filename, table_index, include_formatting)

    @mcp.tool(annotations=ToolAnnotations(title="Set Table Cell", destructiveHint=True))
    async def word_live_set_table_cell(
        filename: str = None,
        table_index: int = 1,
        row: int = 1,
        col: int = 1,
        text: str = "",
        bold: bool = None,
        italic: bool = None,
        font_name: str = None,
        font_size: float = None,
        font_color: str = None,
        alignment: str = None,
    ):
        """Set text and formatting of a single table cell."""
        return await table_com_tools.word_live_set_table_cell(
            filename, table_index, row, col, text, bold, italic, font_name, font_size, font_color, alignment,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Set Table Data", destructiveHint=True))
    async def word_live_set_table_data(filename: str = None, table_index: int = 1, data: list = None, start_row: int = 1, start_col: int = 1):
        """Write a 2D array of values into a table. data = [['r1c1','r1c2'],['r2c1','r2c2']]."""
        return await table_com_tools.word_live_set_table_data(filename, table_index, data, start_row, start_col)

    @mcp.tool(annotations=ToolAnnotations(title="Insert Table Row", destructiveHint=True))
    async def word_live_insert_table_row(filename: str = None, table_index: int = 1, before_row: int = None, count: int = 1):
        """Insert rows into a table. before_row=None appends at end."""
        return await table_com_tools.word_live_insert_table_row(filename, table_index, before_row, count)

    @mcp.tool(annotations=ToolAnnotations(title="Delete Table Row", destructiveHint=True))
    async def word_live_delete_table_row(filename: str = None, table_index: int = 1, row: int = 1, count: int = 1):
        """Delete one or more rows from a table."""
        return await table_com_tools.word_live_delete_table_row(filename, table_index, row, count)

    @mcp.tool(annotations=ToolAnnotations(title="Insert Table Column", destructiveHint=True))
    async def word_live_insert_table_col(filename: str = None, table_index: int = 1, before_col: int = None, count: int = 1):
        """Insert columns into a table. before_col=None appends at end."""
        return await table_com_tools.word_live_insert_table_col(filename, table_index, before_col, count)

    @mcp.tool(annotations=ToolAnnotations(title="Delete Table Column", destructiveHint=True))
    async def word_live_delete_table_col(filename: str = None, table_index: int = 1, col: int = 1, count: int = 1):
        """Delete one or more columns from a table."""
        return await table_com_tools.word_live_delete_table_col(filename, table_index, col, count)

    @mcp.tool(annotations=ToolAnnotations(title="Delete Table", destructiveHint=True))
    async def word_live_delete_table(filename: str = None, table_index: int = 1):
        """Delete an entire table from the document."""
        return await table_com_tools.word_live_delete_table(filename, table_index)

    @mcp.tool(annotations=ToolAnnotations(title="Set Table Row Height", destructiveHint=True))
    async def word_live_set_table_row_height(filename: str = None, table_index: int = 1, row: int = 1, height_cm: float = 1.0, exact: bool = False):
        """Set table row height in centimetres. exact=True clips content to exact height."""
        return await table_com_tools.word_live_set_table_row_height(filename, table_index, row, height_cm, exact)

    @mcp.tool(annotations=ToolAnnotations(title="Set Table Column Width", destructiveHint=True))
    async def word_live_set_table_col_width(filename: str = None, table_index: int = 1, col: int = 1, width_cm: float = 3.0):
        """Set table column width in centimetres."""
        return await table_com_tools.word_live_set_table_col_width(filename, table_index, col, width_cm)

    @mcp.tool(annotations=ToolAnnotations(title="Sort Table", destructiveHint=True))
    async def word_live_sort_table(filename: str = None, table_index: int = 1, sort_col: int = 1, ascending: bool = True, has_header: bool = True):
        """Sort a table by a specified column."""
        return await table_com_tools.word_live_sort_table(filename, table_index, sort_col, ascending, has_header)

    # ── VBA / macro tools ─────────────────────────────────────────────────────

    @mcp.tool(annotations=ToolAnnotations(title="List Macros", readOnlyHint=True))
    async def word_live_list_macros(filename: str = None):
        """List all VBA Sub/Function procedures available in the document and Normal.dotm."""
        return await vba_tools.word_live_list_macros(filename)

    @mcp.tool(annotations=ToolAnnotations(title="Run Macro", destructiveHint=True))
    async def word_live_run_macro(macro_name: str, filename: str = None, args: list = None):
        """Run a VBA macro by name. Use word_live_list_macros to discover available macros."""
        return await vba_tools.word_live_run_macro(macro_name, filename, args)

    @mcp.tool(annotations=ToolAnnotations(title="Get Macro Code", readOnlyHint=True))
    async def word_live_get_macro_code(module_name: str, filename: str = None, scope: str = "document"):
        """Get VBA source code of a module. scope: 'document' or 'normal'."""
        return await vba_tools.word_live_get_macro_code(module_name, filename, scope)

    # ── Snapshot undo ─────────────────────────────────────────────────────────

    @mcp.tool(annotations=ToolAnnotations(title="Take Snapshot (for Undo)", destructiveHint=False))
    async def word_take_snapshot(filename: str = None):
        """Save a document snapshot before making changes. Call word_undo_last to restore it."""
        return await undo_tools.word_take_snapshot(filename)

    @mcp.tool(annotations=ToolAnnotations(title="Undo Last (Restore Snapshot)", destructiveHint=True))
    async def word_undo_last(filename: str = None):
        """Restore the document to its most recent snapshot state, discarding changes since then."""
        return await undo_tools.word_undo_last(filename)

    @mcp.tool(annotations=ToolAnnotations(title="List Snapshots", readOnlyHint=True))
    async def word_list_snapshots(filename: str = None):
        """List available document snapshots for undo."""
        return await undo_tools.word_list_snapshots(filename)

    @mcp.tool(annotations=ToolAnnotations(title="Clear Snapshots", destructiveHint=True))
    async def word_clear_snapshots(filename: str = None):
        """Delete all snapshots for a document (or all documents)."""
        return await undo_tools.word_clear_snapshots(filename)

    # ── Multi-column layout ───────────────────────────────────────────────────

    @mcp.tool(annotations=ToolAnnotations(title="Set Column Layout", destructiveHint=True))
    async def word_live_set_column_layout(
        filename: str = None,
        section_index: int = 1,
        columns: int = 2,
        equal_width: bool = True,
        space_between_cm: float = 1.27,
        line_between: bool = False,
        column_widths_cm: list = None,
        space_after_cm: list = None,
    ):
        """Set multi-column layout for a section (1–9 columns, equal or custom widths)."""
        return await column_layout_tools.word_live_set_column_layout(
            filename, section_index, columns, equal_width,
            space_between_cm, line_between, column_widths_cm, space_after_cm,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Get Column Layout", readOnlyHint=True))
    async def word_live_get_column_layout(filename: str = None, section_index: int = 1):
        """Get the column layout (count, widths, spacing, line-between) of a section."""
        return await column_layout_tools.word_live_get_column_layout(filename, section_index)

    @mcp.tool(annotations=ToolAnnotations(title="Insert Column Break", destructiveHint=True))
    async def word_live_insert_column_break(filename: str = None, paragraph_index: int = None):
        """Insert a column break after a paragraph, forcing following text to the next column."""
        return await column_layout_tools.word_live_insert_column_break(filename, paragraph_index)

    # ── Field code management ─────────────────────────────────────────────────

    @mcp.tool(annotations=ToolAnnotations(title="List Fields", readOnlyHint=True))
    async def word_live_list_fields(filename: str = None):
        """List all field codes in the document with their type, code, and current result text."""
        return await field_tools.word_live_list_fields(filename)

    @mcp.tool(annotations=ToolAnnotations(title="Insert Field", destructiveHint=True))
    async def word_live_insert_field(
        filename: str = None,
        field_type: str = "date",
        paragraph_index: int = None,
        position: str = "end",
        field_code: str = None,
        preserve_formatting: bool = True,
    ):
        """Insert a field code (date, page, num_pages, filename, author, title, seq, custom, etc.)."""
        return await field_tools.word_live_insert_field(
            filename, field_type, paragraph_index, position, field_code, preserve_formatting,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Update Fields", destructiveHint=True))
    async def word_live_update_fields(filename: str = None, field_index: int = None):
        """Recalculate fields. field_index=None updates all; otherwise updates just that field."""
        return await field_tools.word_live_update_fields(filename, field_index)

    @mcp.tool(annotations=ToolAnnotations(title="Delete Field", destructiveHint=True))
    async def word_live_delete_field(
        filename: str = None,
        field_index: int = 1,
        replace_with_result: bool = True,
    ):
        """Delete a field. replace_with_result=True keeps the display text as plain text."""
        return await field_tools.word_live_delete_field(filename, field_index, replace_with_result)

    @mcp.tool(annotations=ToolAnnotations(title="Update TOC", destructiveHint=True))
    async def word_live_update_toc(
        filename: str = None,
        toc_index: int = 1,
        update_page_numbers_only: bool = False,
    ):
        """Refresh a Table of Contents after headings were added, removed, or renamed."""
        return await field_tools.word_live_update_toc(filename, toc_index, update_page_numbers_only)

    @mcp.tool(annotations=ToolAnnotations(title="Update All TOCs", destructiveHint=True))
    async def word_live_update_all_tocs(filename: str = None):
        """Refresh all Tables of Contents in the document at once."""
        return await field_tools.word_live_update_all_tocs(filename)

    # ── Mail merge ────────────────────────────────────────────────────────────

    @mcp.tool(annotations=ToolAnnotations(title="Setup Mail Merge", destructiveHint=True))
    async def word_live_setup_mail_merge(
        filename: str = None,
        data_source_path: str = None,
        merge_type: str = "letters",
        header_row: bool = True,
    ):
        """Connect a document to a mail merge data source (Excel, CSV, Access, etc.)."""
        return await mail_merge_tools.word_live_setup_mail_merge(
            filename, data_source_path, merge_type, header_row,
        )

    @mcp.tool(annotations=ToolAnnotations(title="List Merge Fields", readOnlyHint=True))
    async def word_live_list_merge_fields(filename: str = None):
        """List data field names from the connected mail merge data source."""
        return await mail_merge_tools.word_live_list_merge_fields(filename)

    @mcp.tool(annotations=ToolAnnotations(title="Insert Merge Field", destructiveHint=True))
    async def word_live_insert_merge_field(
        filename: str = None,
        field_name: str = "",
        paragraph_index: int = None,
        position: str = "end",
    ):
        """Insert a «MergeField» placeholder into a mail merge template."""
        return await mail_merge_tools.word_live_insert_merge_field(
            filename, field_name, paragraph_index, position,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Preview Merge Record", destructiveHint=False))
    async def word_live_preview_merge(filename: str = None, record_index: int = 1):
        """Preview how a specific data record looks merged into the template. 0 = back to template."""
        return await mail_merge_tools.word_live_preview_merge(filename, record_index)

    @mcp.tool(annotations=ToolAnnotations(title="Execute Mail Merge", destructiveHint=True))
    async def word_live_execute_merge(
        filename: str = None,
        output_path: str = None,
        records: str = "all",
        from_record: int = 1,
        to_record: int = None,
        merge_to: str = "new_document",
        email_address_field: str = None,
        email_subject: str = None,
    ):
        """Execute the mail merge: produces merged output as new document, printer, or email."""
        return await mail_merge_tools.word_live_execute_merge(
            filename, output_path, records, from_record, to_record,
            merge_to, email_address_field, email_subject,
        )

    # ── Image formatting ──────────────────────────────────────────────────────

    @mcp.tool(annotations=ToolAnnotations(title="Set Image Wrap", destructiveHint=True))
    async def word_live_set_image_wrap(
        filename: str = None,
        shape_name: str = None,
        shape_index: int = None,
        wrap_type: str = "square",
        wrap_side: str = "both",
    ):
        """Set text wrapping style for a floating shape: square, tight, through, top_bottom, behind, in_front."""
        return await image_format_tools.word_live_set_image_wrap(
            filename, shape_name, shape_index, wrap_type, wrap_side,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Set Image Position", destructiveHint=True))
    async def word_live_set_image_position(
        filename: str = None,
        shape_name: str = None,
        shape_index: int = None,
        horizontal_position_cm: float = None,
        vertical_position_cm: float = None,
        horizontal_relative_to: str = "page",
        vertical_relative_to: str = "page",
    ):
        """Set the absolute position of a floating image relative to page, margin, or column."""
        return await image_format_tools.word_live_set_image_position(
            filename, shape_name, shape_index,
            horizontal_position_cm, vertical_position_cm,
            horizontal_relative_to, vertical_relative_to,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Crop Image", destructiveHint=True))
    async def word_live_crop_image(
        filename: str = None,
        image_index: int = 1,
        top_cm: float = 0.0,
        bottom_cm: float = 0.0,
        left_cm: float = 0.0,
        right_cm: float = 0.0,
    ):
        """Crop an inline image by specifying how much to remove from each side in cm."""
        return await image_format_tools.word_live_crop_image(
            filename, image_index, top_cm, bottom_cm, left_cm, right_cm,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Set Image Brightness/Contrast", destructiveHint=True))
    async def word_live_set_image_brightness_contrast(
        filename: str = None,
        image_index: int = 1,
        brightness: float = 0.0,
        contrast: float = 0.0,
    ):
        """Adjust brightness (-1.0 to 1.0) and contrast (-1.0 to 1.0) of an inline image."""
        return await image_format_tools.word_live_set_image_brightness_contrast(
            filename, image_index, brightness, contrast,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Convert Image to Floating", destructiveHint=True))
    async def word_live_convert_to_floating(
        filename: str = None,
        image_index: int = 1,
        wrap_type: str = "square",
        width_cm: float = None,
        height_cm: float = None,
    ):
        """Convert an inline image to a floating shape so it can be freely positioned."""
        return await image_format_tools.word_live_convert_to_floating(
            filename, image_index, wrap_type, width_cm, height_cm,
        )

    # ── Advanced find & replace ───────────────────────────────────────────────

    @mcp.tool(annotations=ToolAnnotations(title="Advanced Find & Replace", destructiveHint=True))
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
    ):
        """Advanced find & replace with wildcards, whole-word, case matching, and formatting constraints."""
        return await find_replace_tools.word_live_find_replace_advanced(
            filename, find_text, replace_text, match_case, whole_word,
            use_wildcards, sounds_like, all_word_forms, replace_all,
            find_font_name, find_font_size, find_bold, find_italic, find_style,
            replace_font_name, replace_font_size, replace_bold, replace_italic, replace_style,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Find All Occurrences", readOnlyHint=True))
    async def word_live_find_all_occurrences(
        filename: str = None,
        find_text: str = "",
        match_case: bool = False,
        whole_word: bool = False,
        use_wildcards: bool = False,
        max_results: int = 100,
    ):
        """Find all occurrences of text and return their character positions, page numbers, and paragraph indices."""
        return await find_replace_tools.word_live_find_all_occurrences(
            filename, find_text, match_case, whole_word, use_wildcards, max_results,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Get Document Statistics", readOnlyHint=True))
    async def word_live_get_document_statistics(filename: str = None):
        """Get word count, character count, page count, paragraph count, and counts of tables/images/comments."""
        return await find_replace_tools.word_live_get_document_statistics(filename)

    # ── Navigation ────────────────────────────────────────────────────────────

    @mcp.tool(annotations=ToolAnnotations(title="Go To Page", destructiveHint=False))
    async def word_live_go_to_page(filename: str = None, page_number: int = 1):
        """Navigate the Word window to a specific page number."""
        return await navigation_tools.word_live_go_to_page(filename, page_number)

    @mcp.tool(annotations=ToolAnnotations(title="Go To Heading", destructiveHint=False))
    async def word_live_go_to_heading(
        filename: str = None,
        heading_index: int = None,
        heading_text: str = None,
    ):
        """Navigate to a heading by its order (heading_index) or partial text match."""
        return await navigation_tools.word_live_go_to_heading(filename, heading_index, heading_text)

    @mcp.tool(annotations=ToolAnnotations(title="Go To Bookmark", destructiveHint=False))
    async def word_live_go_to_bookmark(filename: str = None, bookmark_name: str = ""):
        """Navigate the Word window to a named bookmark."""
        return await navigation_tools.word_live_go_to_bookmark(filename, bookmark_name)

    @mcp.tool(annotations=ToolAnnotations(title="Scroll To Paragraph", destructiveHint=False))
    async def word_live_scroll_to_paragraph(filename: str = None, paragraph_index: int = 1):
        """Scroll the Word view so the specified paragraph is visible."""
        return await navigation_tools.word_live_scroll_to_paragraph(filename, paragraph_index)

    @mcp.tool(annotations=ToolAnnotations(title="List Bookmarks", readOnlyHint=True))
    async def word_live_list_bookmarks(filename: str = None):
        """List all bookmarks in the document with names, positions, and page numbers."""
        return await navigation_tools.word_live_list_bookmarks(filename)

    # ── List formatting ───────────────────────────────────────────────────────

    @mcp.tool(annotations=ToolAnnotations(title="Restart List Numbering", destructiveHint=True))
    async def word_live_restart_list_numbering(filename: str = None, paragraph_index: int = 1):
        """Restart a numbered list at a paragraph so it counts from 1."""
        return await list_tools.word_live_restart_list_numbering(filename, paragraph_index)

    @mcp.tool(annotations=ToolAnnotations(title="Set List Level", destructiveHint=True))
    async def word_live_set_list_level(
        filename: str = None,
        paragraph_index: int = 1,
        start_paragraph: int = None,
        end_paragraph: int = None,
        level: int = 1,
    ):
        """Promote or demote list items to level 1–9."""
        return await list_tools.word_live_set_list_level(
            filename, paragraph_index, start_paragraph, end_paragraph, level,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Continue List Numbering", destructiveHint=True))
    async def word_live_continue_list_numbering(filename: str = None, paragraph_index: int = 1):
        """Make a paragraph continue numbering from the previous list (un-restart)."""
        return await list_tools.word_live_continue_list_numbering(filename, paragraph_index)

    @mcp.tool(annotations=ToolAnnotations(title="Set List Start Value", destructiveHint=True))
    async def word_live_set_list_start_value(
        filename: str = None,
        paragraph_index: int = 1,
        start_value: int = 1,
    ):
        """Set a custom starting number for a numbered list item."""
        return await list_tools.word_live_set_list_start_value(filename, paragraph_index, start_value)

    @mcp.tool(annotations=ToolAnnotations(title="Apply List Style", destructiveHint=True))
    async def word_live_apply_list_style(
        filename: str = None,
        paragraph_index: int = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        list_type: str = "bullet",
        level: int = 1,
    ):
        """Apply bullet/number/roman/alpha list style to paragraphs, or remove list formatting."""
        return await list_tools.word_live_apply_list_style(
            filename, paragraph_index, start_paragraph, end_paragraph, list_type, level,
        )

    # ── Document diff ─────────────────────────────────────────────────────────

    @mcp.tool(annotations=ToolAnnotations(title="Word Diff", readOnlyHint=True))
    def word_diff(path1: str, path2: str, max_diffs: int = 200):
        """Compare two Word (.docx) files and return paragraph- and table-level differences.

        Works cross-platform with no Word installation required. Useful for reviewing
        what changed between document versions or before/after AI edits.

        Args:
            path1: Path to the "before" document.
            path2: Path to the "after" document.
            max_diffs: Maximum number of differences to return (default 200).

        Returns JSON with changed_paragraphs, added_paragraphs, removed_paragraphs,
        changed_tables, truncated flag, and a summary of counts.
        """
        return word_diff_tools.word_diff(path1, path2, max_diffs)


def run_server():
    """Run the Word Document MCP Server with configurable transport."""
    # Get transport configuration
    config = get_transport_config()
    
    # Setup logging
    # setup_logging(config['debug'])
    
    # Monkey-patch Document.save() to preserve comments.xml and other custom parts
    from word_document_server.utils.save_utils import install_save_hook
    install_save_hook()

    # Monkey-patch PhysPkgReader to detect Word-locked files
    from word_document_server.utils.path_utils import install_path_hook
    install_path_hook()

    # Register all tools
    register_tools()
    
    # Print startup information
    transport_type = config['transport']
    print(f"Starting Word Document MCP Server with {transport_type} transport...", file=sys.stderr)
    
    # if config['debug']:
    #     print(f"Configuration: {config}")
    
    try:
        if transport_type == 'stdio':
            # Run with stdio transport (default, backward compatible)
            print("Server running on stdio transport", file=sys.stderr)
            mcp.run(transport='stdio')
            
        elif transport_type == 'streamable-http':
            # Run with streamable HTTP transport
            print(f"Server running on streamable-http transport at http://{config['host']}:{config['port']}{config['path']}", file=sys.stderr)
            mcp.run(
                transport='streamable-http',
                host=config['host'],
                port=config['port'],
                path=config['path']
            )
            
        elif transport_type == 'sse':
            # Run with SSE transport
            print(f"Server running on SSE transport at http://{config['host']}:{config['port']}{config['sse_path']}", file=sys.stderr)
            mcp.run(
                transport='sse',
                host=config['host'],
                port=config['port'],
                path=config['sse_path']
            )
            
    except KeyboardInterrupt:
        print("\nShutting down server...", file=sys.stderr)
    except Exception as e:
        print(f"Error starting server: {e}", file=sys.stderr)
        if config['debug']:
            import traceback
            traceback.print_exc()
        sys.exit(1)
    
    return mcp


def main():
    """Main entry point for the server."""
    run_server()


if __name__ == "__main__":
    main()
