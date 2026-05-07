"""Enhanced table management tools for open Word documents via COM.

Covers: data read/write, row/column insert/delete, merge, sort, delete table,
row height, column width, and cell formatting.
"""

import json
import sys

_PTS_PER_CM = 28.35


async def word_live_get_table_data(
    filename: str = None,
    table_index: int = 1,
    include_formatting: bool = False,
) -> str:
    """Read all cell values from a table in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        table_index: 1-based table index (first table = 1).
        include_formatting: If True, also return basic font info per cell.

    Returns:
        JSON with 2D array of cell text, table dimensions, and optional formatting.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Table COM tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.Tables.Count
        if table_index < 1 or table_index > total:
            return json.dumps({"error": f"table_index {table_index} out of range (1-{total})"})

        table = doc.Tables(table_index)
        rows = table.Rows.Count
        cols = table.Columns.Count

        data = []
        for r in range(1, rows + 1):
            row_data = []
            for c in range(1, cols + 1):
                try:
                    cell = table.Cell(r, c)
                    text = cell.Range.Text or ""
                    # Remove the cell-end marker \x07\r
                    text = text.rstrip("\x07\r\n")
                    if include_formatting:
                        font = cell.Range.Font
                        row_data.append({
                            "text": text,
                            "bold": font.Bold if font.Bold != 9999999 else None,
                            "italic": font.Italic if font.Italic != 9999999 else None,
                            "font_name": font.Name if font.Name != 9999999 else None,
                            "font_size": font.Size if font.Size != 9999999 else None,
                        })
                    else:
                        row_data.append(text)
                except Exception:
                    row_data.append("") if not include_formatting else row_data.append({"text": ""})
            data.append(row_data)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "table_index": table_index,
            "rows": rows,
            "cols": cols,
            "data": data,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


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
) -> str:
    """Set the text and formatting of a single table cell in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        table_index: 1-based table index.
        row: 1-based row index.
        col: 1-based column index.
        text: New cell content.
        bold / italic: Font weight/style.
        font_name: Font family.
        font_size: Font size in points.
        font_color: Text color as "#RRGGBB".
        alignment: Cell text alignment — "left", "center", "right", "justify".

    Returns:
        JSON confirming the update.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Table COM tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        ALIGN_MAP = {"left": 0, "center": 1, "right": 2, "justify": 3}
        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.Tables.Count
        if table_index < 1 or table_index > total:
            return json.dumps({"error": f"table_index {table_index} out of range (1-{total})"})

        table = doc.Tables(table_index)
        cell = table.Cell(row, col)

        with undo_record(app, "MCP: Set Table Cell"):
            cell.Range.Text = text
            font = cell.Range.Font
            if bold is not None:
                font.Bold = bold
            if italic is not None:
                font.Italic = italic
            if font_name:
                font.Name = font_name
            if font_size:
                font.Size = font_size
            if font_color:
                c = font_color.lstrip("#")
                font.Color = int(c[0:2], 16) + (int(c[2:4], 16) << 8) + (int(c[4:6], 16) << 16)
            if alignment and alignment in ALIGN_MAP:
                for para in cell.Range.Paragraphs:
                    para.Format.Alignment = ALIGN_MAP[alignment]

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "table_index": table_index,
            "row": row,
            "col": col,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_set_table_data(
    filename: str = None,
    table_index: int = 1,
    data: list = None,
    start_row: int = 1,
    start_col: int = 1,
) -> str:
    """Write a 2D array of values into a table in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        table_index: 1-based table index.
        data: 2D list of strings. e.g. [["Name","Age"],["Alice","30"]].
        start_row: 1-based row to start writing at (default 1).
        start_col: 1-based column to start writing at (default 1).

    Returns:
        JSON with count of cells written.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Table COM tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.Tables.Count
        if table_index < 1 or table_index > total:
            return json.dumps({"error": f"table_index {table_index} out of range (1-{total})"})

        table = doc.Tables(table_index)
        rows_count = table.Rows.Count
        cols_count = table.Columns.Count
        cells_written = 0

        with undo_record(app, "MCP: Set Table Data"):
            for ri, row_data in enumerate(data or []):
                r = start_row + ri
                if r > rows_count:
                    break
                for ci, cell_text in enumerate(row_data):
                    c = start_col + ci
                    if c > cols_count:
                        break
                    try:
                        table.Cell(r, c).Range.Text = str(cell_text)
                        cells_written += 1
                    except Exception:
                        pass

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "table_index": table_index,
            "cells_written": cells_written,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_insert_table_row(
    filename: str = None,
    table_index: int = 1,
    before_row: int = None,
    count: int = 1,
) -> str:
    """Insert one or more rows into a table in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        table_index: 1-based table index.
        before_row: Insert before this row (1-based). None = append at end.
        count: Number of rows to insert (default 1).

    Returns:
        JSON with updated row count.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Table COM tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.Tables.Count
        if table_index < 1 or table_index > total:
            return json.dumps({"error": f"table_index {table_index} out of range (1-{total})"})

        table = doc.Tables(table_index)
        with undo_record(app, "MCP: Insert Table Row"):
            for _ in range(count):
                if before_row is not None:
                    current_rows = table.Rows.Count
                    if before_row < 1 or before_row > current_rows:
                        table.Rows.Add()
                    else:
                        table.Rows.Add(BeforeRow=table.Rows(before_row))
                else:
                    table.Rows.Add()

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "table_index": table_index,
            "rows_now": table.Rows.Count,
            "inserted": count,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_delete_table_row(
    filename: str = None,
    table_index: int = 1,
    row: int = 1,
    count: int = 1,
) -> str:
    """Delete one or more rows from a table in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        table_index: 1-based table index.
        row: 1-based row to delete (first row to delete if count > 1).
        count: Number of consecutive rows to delete (default 1).

    Returns:
        JSON with updated row count.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Table COM tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.Tables.Count
        if table_index < 1 or table_index > total:
            return json.dumps({"error": f"table_index {table_index} out of range (1-{total})"})

        table = doc.Tables(table_index)
        current = table.Rows.Count
        if row < 1 or row > current:
            return json.dumps({"error": f"row {row} out of range (1-{current})"})

        with undo_record(app, "MCP: Delete Table Row"):
            for _ in range(min(count, table.Rows.Count - row + 1)):
                table.Rows(row).Delete()

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "table_index": table_index,
            "rows_now": table.Rows.Count,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_insert_table_col(
    filename: str = None,
    table_index: int = 1,
    before_col: int = None,
    count: int = 1,
) -> str:
    """Insert one or more columns into a table in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        table_index: 1-based table index.
        before_col: Insert before this column (1-based). None = append at end.
        count: Number of columns to insert (default 1).

    Returns:
        JSON with updated column count.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Table COM tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.Tables.Count
        if table_index < 1 or table_index > total:
            return json.dumps({"error": f"table_index {table_index} out of range (1-{total})"})

        table = doc.Tables(table_index)
        with undo_record(app, "MCP: Insert Table Column"):
            for _ in range(count):
                if before_col is not None:
                    cur_cols = table.Columns.Count
                    if before_col < 1 or before_col > cur_cols:
                        table.Columns.Add()
                    else:
                        table.Columns.Add(BeforeColumn=table.Columns(before_col))
                else:
                    table.Columns.Add()

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "table_index": table_index,
            "cols_now": table.Columns.Count,
            "inserted": count,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_delete_table_col(
    filename: str = None,
    table_index: int = 1,
    col: int = 1,
    count: int = 1,
) -> str:
    """Delete one or more columns from a table in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        table_index: 1-based table index.
        col: 1-based column to delete.
        count: Number of consecutive columns to delete (default 1).

    Returns:
        JSON with updated column count.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Table COM tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.Tables.Count
        if table_index < 1 or table_index > total:
            return json.dumps({"error": f"table_index {table_index} out of range (1-{total})"})

        table = doc.Tables(table_index)
        cur = table.Columns.Count
        if col < 1 or col > cur:
            return json.dumps({"error": f"col {col} out of range (1-{cur})"})

        with undo_record(app, "MCP: Delete Table Column"):
            for _ in range(min(count, table.Columns.Count - col + 1)):
                table.Columns(col).Delete()

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "table_index": table_index,
            "cols_now": table.Columns.Count,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_delete_table(
    filename: str = None,
    table_index: int = 1,
) -> str:
    """Delete an entire table from an open Word document.

    Args:
        filename: Document name or path (None = active document).
        table_index: 1-based table index.

    Returns:
        JSON confirming deletion.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Table COM tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.Tables.Count
        if table_index < 1 or table_index > total:
            return json.dumps({"error": f"table_index {table_index} out of range (1-{total})"})

        with undo_record(app, "MCP: Delete Table"):
            doc.Tables(table_index).Delete()

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "deleted_table_index": table_index,
            "tables_remaining": doc.Tables.Count,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_set_table_row_height(
    filename: str = None,
    table_index: int = 1,
    row: int = 1,
    height_cm: float = 1.0,
    exact: bool = False,
) -> str:
    """Set the height of a table row in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        table_index: 1-based table index.
        row: 1-based row index.
        height_cm: Row height in centimetres.
        exact: True = exact height (clips content). False = at-least height (default).

    Returns:
        JSON confirming update.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Table COM tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.Tables.Count
        if table_index < 1 or table_index > total:
            return json.dumps({"error": f"table_index {table_index} out of range (1-{total})"})

        table = doc.Tables(table_index)
        tr = table.Rows(row)

        with undo_record(app, "MCP: Set Row Height"):
            tr.HeightRule = 2 if exact else 1   # wdRowHeightExactly=2, wdRowHeightAtLeast=1
            tr.Height = height_cm * _PTS_PER_CM

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "table_index": table_index,
            "row": row,
            "height_cm": height_cm,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_set_table_col_width(
    filename: str = None,
    table_index: int = 1,
    col: int = 1,
    width_cm: float = 3.0,
) -> str:
    """Set the width of a table column in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        table_index: 1-based table index.
        col: 1-based column index.
        width_cm: Column width in centimetres.

    Returns:
        JSON confirming update.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Table COM tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.Tables.Count
        if table_index < 1 or table_index > total:
            return json.dumps({"error": f"table_index {table_index} out of range (1-{total})"})

        with undo_record(app, "MCP: Set Column Width"):
            doc.Tables(table_index).Columns(col).Width = width_cm * _PTS_PER_CM

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "table_index": table_index,
            "col": col,
            "width_cm": width_cm,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_sort_table(
    filename: str = None,
    table_index: int = 1,
    sort_col: int = 1,
    ascending: bool = True,
    has_header: bool = True,
) -> str:
    """Sort a table by a specified column in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        table_index: 1-based table index.
        sort_col: Column to sort by (1-based).
        ascending: Sort ascending (default True). False = descending.
        has_header: Treat first row as header, exclude from sort (default True).

    Returns:
        JSON confirming the sort.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Table COM tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.Tables.Count
        if table_index < 1 or table_index > total:
            return json.dumps({"error": f"table_index {table_index} out of range (1-{total})"})

        table = doc.Tables(table_index)
        with undo_record(app, "MCP: Sort Table"):
            table.Sort(
                ExcludeHeader=has_header,
                FieldNumber=f"Column {sort_col}",
                SortOrder=0 if ascending else 1,   # wdSortOrderAscending=0, wdSortOrderDescending=1
                SortFieldType=0,   # wdSortFieldAlphanumeric
            )

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "table_index": table_index,
            "sort_col": sort_col,
            "ascending": ascending,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})
