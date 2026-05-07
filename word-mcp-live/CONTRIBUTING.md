# Contributing to word-mcp-live

Thanks for your interest in contributing! This guide covers the basics of setting up a development environment and adding new tools.

## Development Setup

```bash
git clone https://github.com/ykarapazar/word-mcp-live.git
cd word-mcp-live
pip install -e ".[dev]"
```

For Windows Live tools, you also need:
- Windows 10/11 with Microsoft Word installed
- `pip install pywin32`

## Project Structure

```
word_document_server/
  main.py               # FastMCP server — all tools registered here
  defaults.py           # Default author name, initials
  tools/
    document_tools.py   # Document management (create, copy, info)
    content_tools.py    # Content manipulation (paragraphs, tables, images)
    format_tools.py     # Formatting (text, tables, styles)
    layout_tools.py     # Page layout, headers, watermarks
    footnote_tools.py   # Footnotes and endnotes
    protection_tools.py # Password protection, signatures
    comment_tools.py    # Read comments
    comment_write_tools.py  # Write comments
    hyperlink_tools.py  # Hyperlink management
    tracked_changes_tools.py  # Tracked changes via OOXML
    live_tools.py       # Windows COM editing tools
    live_read_tools.py  # Windows COM reading tools
    live_layout_tools.py    # Windows COM layout tools
    screen_capture_tools.py # Word window screenshot
  core/
    word_com.py         # COM helpers (get_word_app, undo_record)
```

## Code Style

- Python 3.11+ — use modern syntax (type hints, `match` statements where appropriate)
- Every tool function must have a docstring — it becomes the tool's description in MCP
- Use `description=impl_func.__doc__` in `@mcp.tool()` to keep a single source of truth
- Follow existing patterns in `main.py` for tool registration

## Adding a New Tool

### 1. Write the implementation

Add your function to the appropriate file in `word_document_server/tools/`. For a new cross-platform tool:

```python
# word_document_server/tools/content_tools.py

def my_new_tool(file_path: str, param: str) -> str:
    """Short description of what this tool does.

    Args:
        file_path: Path to the Word document
        param: What this parameter controls

    Returns:
        Success message or result
    """
    doc = Document(file_path)
    # ... implementation ...
    doc.save(file_path)
    return f"Done: {param}"
```

### 2. Register in main.py

```python
from word_document_server.tools import content_tools

@mcp.tool(
    annotations=ToolAnnotations(readOnlyHint=False),
    description=content_tools.my_new_tool.__doc__,
)
def my_new_tool(file_path: str, param: str) -> str:
    return content_tools.my_new_tool(file_path, param)
```

### 3. For Windows Live tools

Live tools use COM automation via `pywin32`. Use the helpers in `core/word_com.py`:

```python
from word_document_server.core.word_com import get_word_app, find_open_document, undo_record

def word_live_my_tool(filename: str = None) -> str:
    """Description of the live tool."""
    app = get_word_app()
    doc = find_open_document(app, filename)

    with undo_record(app, "MCP: My Tool"):
        # ... COM operations ...
        pass

    return "Done"
```

All destructive live tools must be wrapped with `undo_record` so each operation appears as a single Ctrl+Z entry in Word.

## Running Tests

```bash
pytest tests/
```

## Pull Request Guidelines

1. Keep PRs focused — one feature or fix per PR
2. Update the tool count in `README.md` badges and text if you add/remove tools
3. Add a changelog entry under `## [Unreleased]` in `CHANGELOG.md`
4. Test cross-platform tools on at least one platform; live tools require Windows + Word
