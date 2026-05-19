<div align="center">

[![Install in Cursor](https://cursor.com/deeplink/mcp-install-dark.svg)](https://cursor.com/en/install-mcp?name=word&config=eyJjb21tYW5kIjoidXZ4IiwiYXJncyI6WyJ3b3JkLW1jcC1saXZlIl19)

# word-mcp-live

**The only MCP server that edits Word documents while they're open**

`Native tracked changes` &middot; `Per-action Ctrl+Z` &middot; `Data-loss safeguards` &middot; `Live editing` &middot; `Cross-platform`

[![PyPI](https://img.shields.io/pypi/v/word-mcp-live?color=blue)](https://pypi.org/project/word-mcp-live/)
[![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Platform: Windows + macOS/Linux](https://img.shields.io/badge/platform-Windows%20%2B%20macOS%2FLinux-lightgrey)]()

</div>

---

word-mcp-live gives any AI assistant that supports [MCP](https://modelcontextprotocol.io/) full control of Microsoft Word. Open a document, tell the AI what you need, and watch it happen — formatting, tracked changes, comments, and all. Changes appear live in your open document.

<table>
<tr>
<td width="50%">

### Without word-mcp-live

- AI can discuss your document but can't touch it
- You copy-paste between AI and Word, losing formatting
- Track changes? You do those manually after the fact
- Every edit means save → close → process → reopen

</td>
<td width="50%">

### With word-mcp-live

- "Add a tracked change replacing ABC Corp with XYZ Ltd" — done
- Changes appear live in your open Word document
- Every AI edit is one Ctrl+Z away
- Real tracked changes with your name, not XML hacks

</td>
</tr>
</table>

### See it in action

https://github.com/user-attachments/assets/fbb09af4-1e25-4e49-94d0-45b363278810

## What Sets This Apart

- **Native tracked changes** — Wraps Word's own `Document.TrackRevisions = True`, not lxml/OOXML hacks. Author, timestamp, and reviewer-pane behavior match revisions made by a human editor. Cross-platform fallback exists but is marked deprecated for production — see [CHANGELOG](CHANGELOG.md).
- **Real per-action Ctrl+Z** — Every AI tool call becomes a single entry in Word's undo stack via `Document.UndoRecord.StartCustomRecord`. Made a mistake? One keystroke reverts it.
- **Data-loss safeguards** *(things that have actually broken in production and we fixed)*:
  - **Control-byte filter** — `\x07` (cell separator) passed to find-and-replace previously matched across table cells and could empty an entire document. Now rejected at the boundary with a descriptive error.
  - **Zero-length wildcard guard** — `Find.Execute` with patterns like `*` no longer infinite-loops; 50K replacement ceiling as belt-and-braces.
  - **32K auto-chunking** — Word COM `InsertBefore`/`InsertAfter` silently truncate past ~32K chars; we split and stitch automatically.
  - **Concurrent-COM lock + liveness probe** — Two parallel tool calls no longer race on the COM proxy; a frozen Word (modal dialog) raises a clear error instead of hanging the server.
- **Live editing** — Operates on documents currently open in Word. No save-close-reopen cycle.
- **Threaded comments** *(Windows)* — Add, reply, resolve, and delete comments like a human reviewer.
- **Layout diagnostics** — Scans for `keep_with_next` chains, broken style references, and section-break issues before they become print disasters.
- **Equations & cross-references** — Insert UnicodeMath formulas and auto-updating references to headings, bookmarks, figures, tables.
- **100+ tools** across cross-platform and live modes — see [TOOLS.md](TOOLS.md) for the full list.

## Quick Start

```bash
pip install word-mcp-live
```

Or install from source:

```bash
git clone https://github.com/ykarapazar/word-mcp-live.git
cd word-mcp-live
pip install -e .
```

## Client Installation

<details open>
<summary><b>Claude Desktop</b></summary>

Add to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "word": {
      "command": "uvx",
      "args": ["word-mcp-live"],
      "env": {
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

</details>

<details>
<summary><b>Claude Code</b></summary>

Add to your `.mcp.json`:

```json
{
  "mcpServers": {
    "word": {
      "command": "uvx",
      "args": ["word-mcp-live"],
      "env": {
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

</details>

<details>
<summary><b>Cursor</b></summary>

**One-click:** Click the install button at the top of this page.

**Manual:** Add to `~/.cursor/mcp.json`:

```json
{
  "mcpServers": {
    "word": {
      "command": "uvx",
      "args": ["word-mcp-live"],
      "env": {
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

</details>

<details>
<summary><b>VS Code / Copilot</b></summary>

**One-click:** [Install in VS Code](vscode:mcp/install?%7B%22name%22%3A%20%22word%22%2C%20%22command%22%3A%20%22uvx%22%2C%20%22args%22%3A%20%5B%22word-mcp-live%22%5D%7D)

**Manual:** Add to your VS Code `settings.json`:

```json
{
  "mcp": {
    "servers": {
      "word": {
        "command": "uvx",
        "args": ["word-mcp-live"],
        "env": {
          "MCP_AUTHOR": "Your Name",
          "MCP_AUTHOR_INITIALS": "YN"
        }
      }
    }
  }
}
```

</details>

<details>
<summary><b>Windsurf</b></summary>

Add to `~/.codeium/windsurf/mcp_config.json`:

```json
{
  "mcpServers": {
    "word": {
      "command": "uvx",
      "args": ["word-mcp-live"],
      "env": {
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

</details>

<details>
<summary><b>Docker</b></summary>

```json
{
  "mcpServers": {
    "word": {
      "command": "docker",
      "args": ["run", "-i", "--rm", "ghcr.io/ykarapazar/word-mcp-live"],
      "env": {
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

> Note: Docker mode supports cross-platform tools only. Live editing requires a native Windows install.

</details>

> **`MCP_AUTHOR`** sets your name on tracked changes and comments (default: `"Author"`). **`MCP_AUTHOR_INITIALS`** sets comment initials.

## Two Modes

|  | Works everywhere | Live editing (Word open) |
|---|---|---|
| **What it does** | Create and edit saved .docx files | Edit documents live while you work in Word |
| **Platform** | Windows, macOS, Linux | Windows (COM) and macOS (JXA) |
| **Undo** | File-level saves | Per-action Ctrl+Z (Windows); per-operation undo (macOS) |
| **Best for** | Batch processing, document generation | Interactive editing, formatting, review |

Both modes work together. The AI picks the right one for the task.

### macOS Live Editing (New in v1.5.0)

Live tools now work on macOS via JavaScript for Automation (JXA). Same tool names, same parameters — the server detects your platform and uses the right automation backend.

| Feature | Windows | macOS |
|---------|---------|-------|
| Text read/write/find/replace | COM | JXA |
| Formatting (bold, font, style) | COM | JXA |
| Track changes & revisions | COM | JXA |
| Comments (add, delete, list) | COM | JXA |
| Tables (read, write, add rows) | COM | JXA |
| Page layout, headers, bookmarks | COM | JXA |
| Equations, cross-references | COM | JXA |
| Threaded comment replies | COM | Not available |
| Comment resolve/unresolve | COM | Not available |
| Undo history inspection | COM | Not available |
| Watermarks | COM | Not available |

## Configuration

| Variable | Default | Description |
|----------|---------|-------------|
| `MCP_AUTHOR` | `"Author"` | Author name for tracked changes and comments |
| `MCP_AUTHOR_INITIALS` | `""` | Author initials for comments |
| `MCP_TRANSPORT` | `stdio` | Transport type: `stdio`, `sse`, or `streamable-http` |
| `MCP_HOST` | `0.0.0.0` | Host to bind (for SSE/HTTP transports) |
| `MCP_PORT` | `8000` | Port to bind (for SSE/HTTP transports) |

For remote deployment, see [RENDER_DEPLOYMENT.md](RENDER_DEPLOYMENT.md).

## Example Prompts

Just tell the AI what you want in plain language:

```
"Draft a contract with tracked changes so my colleague can review"
"Format all headings as Cambria 13pt bold and add automatic numbering"
"Add a comment on paragraph 3 asking about the deadline"
"Find every mention of 'ABC Corp' and replace with 'XYZ Ltd' as a tracked change"
"Set the page to A4 landscape with 2cm margins"
"Insert a table of contents based on the document headings"
"Add page numbers in the footer and our company name in the header"
"Insert a cross-reference to Heading 2 in paragraph 5"
```

## Usage Examples

### Example 1: Read a document (cross-platform)

**Tool call:** `get_document_text`
```json
{ "filename": "C:/Documents/report.docx" }
```
**Expected output:**
```json
{
  "status": "success",
  "paragraphs": [
    {"index": 0, "text": "Quarterly Report", "style": "Heading 1"},
    {"index": 1, "text": "Revenue increased by 15% compared to Q3.", "style": "Normal"},
    {"index": 2, "text": "Key Metrics", "style": "Heading 2"}
  ],
  "total_paragraphs": 3
}
```

### Example 2: Live editing with tracked changes (Windows)

**Tool call:** `word_live_replace_text`
```json
{
  "filename": "report.docx",
  "find_text": "ABC Corporation",
  "replace_text": "XYZ Ltd",
  "match_case": true,
  "replace_all": true,
  "track_changes": true
}
```
**Expected output:**
```json
{
  "status": "success",
  "replacements": 4,
  "message": "Replaced 4 occurrences (tracked changes enabled)"
}
```
The replacements appear as tracked changes in Word with strikethrough on "ABC Corporation" and underline on "XYZ Ltd".

### Example 3: Add a comment anchored to text (cross-platform)

**Tool call:** `add_comment`
```json
{
  "filename": "C:/Documents/contract.docx",
  "target_text": "payment within 30 days",
  "comment_text": "Should we extend this to 45 days?",
  "author": "Jane Smith"
}
```
**Expected output:**
```json
{
  "status": "success",
  "message": "Comment added by Jane Smith on 'payment within 30 days'"
}
```
The comment appears in Word's Review panel, anchored to the specified text.

## Workflow Rules for AI Agents

When an LLM agent uses this MCP for document formatting work, two failure modes recur regardless of the underlying model: **silent omissions** ("AI said it changed margins but didn't") and **scope creep** ("AI also rewrote three other things you didn't ask for"). Adding the two rules below to your agent's system prompt or `CLAUDE.md` makes both failure modes structurally hard to hit.

### Rule 1 — Audit-before-done (catches omissions)

Any `.docx` formatting task must end with a `word_audit_against_spec` call whose JSON output is pasted verbatim into the agent's final message. `fail_count > 0` is not "done".

Required workflow:

1. Ask the user for a JSON spec, or convert their natural-language spec into `[{ "id": "<dotted.path>", "expected": <value> }, …]` and confirm.
2. Make edits via `word_live_*` or cross-platform tools.
3. Call `word_live_save` (the audit reads from disk, not from Word's in-memory state).
4. Call `word_audit_against_spec(filename=…, spec_path=…)`.
5. If `fail_count > 0`, follow each item's `fix_hint`, then loop to step 3.
6. Only when `fail_count == 0`, declare completion **and paste the full audit JSON** as evidence.

Boundaries:

- The registered checkers live in `word_document_server/tools/audit_tools.py::CHECKERS`. Spec items the audit doesn't cover must be **explicitly listed** by the agent with the verification method used (PDF render, manual check, `word_diff`) — never silently skipped.
- `unknown_count > 0` means the spec references a checker that isn't registered. Treat it as a fail until the checker is added or the spec is corrected.
- When audit cannot run (non-`.docx` task, missing file), the agent must say so explicitly, not pretend the audit happened.

### Rule 2 — Minimum-change principle (catches scope creep)

The agent does only what was explicitly requested. Scope is made explicit through a three-step contract, not left to the model's judgement.

1. **Before acting**, restate the scope: "I will do: [1, 2, 3]", each item mapped back to a user sentence. Separately list "noticed but not doing" items. When the request is ambiguous, ask before acting.
2. **During each edit**, verify the change maps to a scope item. If it doesn't, stop and ask.
3. **At completion**, deliver in two sections:
   - ✅ **Done** — checked against the scope list
   - ⚠️ **Noticed but not done** — listed for the user to commission separately

Danger signals — when these words appear in your own reasoning, stop and ask:

- "while I'm here / since I'm here / and also" — 90% chance it's scope creep
- "for consistency / to match / for completeness" — not the current task
- "for future-proofing / defensively / for robustness" — preventive expansion the user didn't request

Forbidden by default (unless explicitly requested):

- Editing README / CHANGELOG / unrelated docstrings while working on code
- Renaming variables, refactoring functions, removing "dead" code
- Adding tests, type hints, or error handling
- Fixing unrelated bugs noticed in the same file (list them, don't fix them)
- Expanding a tool's default behavior or parameter set on your own initiative

Scope vs. request specificity:

| Request shape | Default scope |
|---|---|
| "Fix bug X" | **Only** X — don't touch related code |
| "Improve X" | X plus directly related items, listed for review |
| "Review / optimize file Y" | Whole file, but report by category |
| Ambiguous request | **Ask** before acting |

These two rules pair: Rule 1 prevents the agent from claiming completion when items are missing; Rule 2 prevents it from inflating the work beyond what was asked. Both rely on the agent surfacing the *evidence of completion* (audit JSON for Rule 1; scope contract + two-section report for Rule 2) rather than self-reporting "done".

## Tool Reference

**100+ tools** across two modes — see the [complete tool reference](TOOLS.md) for details.

| Category | Count |
|----------|-------|
| Cross-platform (python-docx) | 80 |
| Windows Live (COM automation) | 44 |
| macOS Live (JXA automation) | 40 (of the 44 live tools) |

*Counts are rounded; exact registration list lives in `word_document_server/tool_registry.py`.*

## Requirements

- **Python 3.11+**
- `python-docx`, `fastmcp`, `msoffcrypto-tool` (installed automatically)
- **Windows Live tools:** Windows 10/11 + Microsoft Word + `pywin32` (installed automatically)
- **macOS Live tools:** macOS + Microsoft Word for Mac (uses built-in JXA — no extra dependencies)

> The cross-platform tools work without Word installed — only python-docx is needed.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup, code style, and how to add new tools.

Found a bug? [Open an issue](https://github.com/ykarapazar/word-mcp-live/issues/new?template=bug_report.md).
Have an idea? [Request a feature](https://github.com/ykarapazar/word-mcp-live/issues/new?template=feature_request.md).

## Acknowledgments

Built on top of [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server) by GongRzhe (MIT License).

Additional libraries: [python-docx](https://python-docx.readthedocs.io/) &middot; [FastMCP](https://github.com/modelcontextprotocol/python-sdk) &middot; [pywin32](https://github.com/mhammond/pywin32)

## Privacy

This server runs entirely on your local machine. No data is collected, transmitted, or stored. See the full [Privacy Policy](PRIVACY.md).

## Support

- **Bug reports:** [Open an issue](https://github.com/ykarapazar/word-mcp-live/issues/new?template=bug_report.md)
- **Feature requests:** [Request a feature](https://github.com/ykarapazar/word-mcp-live/issues/new?template=feature_request.md)
- **Discussions:** [GitHub Discussions](https://github.com/ykarapazar/word-mcp-live/discussions)

## License

MIT License — see [LICENSE](LICENSE) for details.

## Star History

<a href="https://star-history.com/#ykarapazar/word-mcp-live&Date">
 <picture>
   <source media="(prefers-color-scheme: dark)" srcset="https://api.star-history.com/svg?repos=ykarapazar/word-mcp-live&type=Date&theme=dark" />
   <source media="(prefers-color-scheme: light)" srcset="https://api.star-history.com/svg?repos=ykarapazar/word-mcp-live&type=Date" />
   <img alt="Star History Chart" src="https://api.star-history.com/svg?repos=ykarapazar/word-mcp-live&type=Date" />
 </picture>
</a>

<!-- mcp-name: io.github.ykarapazar/word-mcp-live -->
