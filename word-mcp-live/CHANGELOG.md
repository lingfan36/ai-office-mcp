# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.8.0] - 2026-05-18

### Added
- **`word_audit_against_spec` — the verification tool that breaks "AI says done" away from "actually done".** Takes a structured JSON spec (list of `{"id": "<dotted.path>", "expected": <value>}` items) and a `.docx`, reads the document via python-docx + lxml, dispatches each id to a registered checker, and returns:
  ```json
  {
    "pass_count": N,
    "fail_count": M,
    "items": [
      { "id": "...", "expected": ..., "actual": ..., "pass": false, "fix_hint": "..." }
    ]
  }
  ```
  Designed to be invoked by the caller (AI or human) immediately before declaring completion — combined with a CLAUDE.md rule that requires the audit JSON to be surfaced before any "done" claim, this physically prevents silent omissions.

- **Initial checker library** (`word_document_server/tools/audit_tools.py::CHECKERS`) covers 25+ id namespaces:
  - `page.*` — margins, gutter, header/footer distance, odd/even pages
  - `body.*` — East Asia font, ASCII font, size, line spacing on body paragraphs
  - `heading1.*` / `heading2.*` — font, size, alignment per heading level
  - `headings.no_level_skip` — detects 1→1.1.1 style skips
  - `patterns.max_consecutive_blank_paragraphs` — caps blank-paragraph runs
  - `patterns.no_forbidden_body_fonts` — blocks Calibri / Arial / 楷体 etc. leakage into body
  - `patterns.no_red_body_text` — flags red runs in body
  - `patterns.no_halfwidth_punct_in_cjk` — catches `.` `,` `;` mixed into CJK paragraphs
  - `patterns.keywords_separator` — verifies 关键词 uses `；` and Keywords uses `;`
  - `references.numbering_format` — every entry starts with `[N]`

- **Extensible by design**: adding a new check is one entry in `CHECKERS` + one function. Each failing item carries an optional `fix_hint` pointing at the MCP tool that would repair it.

- **Demonstrated**: on a freshly generated bad thesis with all 16+ seeded problems, audit catches 23/25 items as FAIL with accurate expected/actual readouts. The 2 PASS items are vacuously-true cases (e.g. no real heading styles → no level-skip possible).

### Recommended CLAUDE.md addition

For users who want to enforce audit-before-done at the agent level, add:

```
## 排版任务完成判定
- 涉及 .docx 排版的任务，完成前必须调用 word_audit_against_spec
- 必须把 audit 返回的 JSON 完整 paste 在最后一条消息里
- fail_count > 0 时不得声明"完成"，必须先修后再 audit
- 调 audit 前先 word_live_save，确保磁盘状态等于编辑状态
```

## [1.7.3] - 2026-05-18

### Added
- **`set_page_layout` / `word_live_set_page_layout` — 5 new properties + cm unit support.** Page setup is the most common task that fails the "spec → checklist" cross-check (header/footer distance from edge, gutter, alternating odd/even headers — all spec-required, none previously settable via MCP, callers had to drop to COM). Both tools now accept:
  - `gutter_cm` / `gutter_inches` — binding margin for double-sided print.
  - `header_distance_cm` / `header_distance_inches` — distance from page edge to header band.
  - `footer_distance_cm` / `footer_distance_inches` — distance from page edge to footer band.
  - `different_first_page` — section-level toggle for cover-page header/footer.
  - `different_odd_and_even_pages` — document-level toggle (cross-platform path writes ``<w:evenAndOddHeaders/>`` to settings.xml; live path uses ``Section.PageSetup.OddAndEvenPagesHeaderFooter``).
- **Every existing dimension param now has a `_cm` twin** (`page_width_cm`, `margin_top_cm`, etc.). If both `_cm` and `_inches` are supplied for the same dimension, cm wins and a note is attached to the response. Removes the cm→in mental arithmetic that AI clients kept botching against Chinese / EU style specs.

### Changed
- **Response now carries `effective` / `effective_pt` dict** — every dimension actually written is echoed back in its canonical unit (cm for cross-platform, points for live). Lets callers diff against spec without re-reading the document. Same pattern as 1.7.2's `effective_line_spacing_pt`.
- **Range validation** — values larger than ~56 inches now raise before touching the document (caught a class of "passing twentieths-of-a-point thinking they're inches" errors).

## [1.7.2] - 2026-05-18

### Fixed
- **`word_live_set_paragraph_spacing` line-spacing semantics** — Word COM's `Paragraph.LineSpacing` is always in points, including when `LineSpacingRule = wdLineSpaceMultiple`. The previous implementation forwarded the raw `line_spacing` argument unchanged, so calling with `line_spacing=1.5, line_spacing_rule="multiple"` produced 1.5-pt fixed line height (text physically overlaps) instead of 1.5× spacing. The tool now:
  - Treats `line_spacing` as a **multiplier** under `rule="multiple"` and converts to points internally as `value × 12`. Values > 10 are kept as-is to preserve any caller that pre-multiplied.
  - **Inferrs the rule** when `line_spacing_rule` is omitted: `value ≤ 10` → multiple (×12 conversion), `value > 10` → exactly (raw points).
  - **Ignores `line_spacing` for preset rules** (`single` / `1.5_lines` / `double`) — Word computes the points itself; previous behavior of overwriting the preset with a stray number is gone.
  - **Sets `LineSpacingRule` before `LineSpacing`** so Word doesn't reinterpret the value during assignment.
  - Returns `effective_line_spacing_rule` and `effective_line_spacing_pt` in the response payload for easy verification.
- macOS JXA `mac_set_paragraph_spacing` still accepts only raw points and does not honor `line_spacing_rule`. Now documented; full parity is deferred.

## [1.7.1] - 2026-05-18

### Fixed
- **All `word_live_*` write operations timing out at 3 s after upgrading to 1.7.0** — the new cross-thread liveness probe in `_com_guard` deadlocked against Word's STA marshalling. The probe ran on a daemon thread but called a property on a COM proxy bound to the main thread's STA; completing that RPC requires the main thread to pump Win32 messages, but the main thread was parked in `Event.wait(timeout)` and never pumped. Every write therefore appeared to "find Word frozen" and returned a misleading "Word is unresponsive" error after `WORD_COM_LIVENESS_TIMEOUT`. Reads bypassed the guard and were unaffected.
  - **Default behaviour change**: liveness probe is now **disabled by default**. Set `WORD_COM_LIVENESS_PROBE=1` to opt in.
  - When opted in, `_probe_with_timeout` now interleaves `pythoncom.PumpWaitingMessages()` while waiting on the worker thread, so the STA marshalling can complete and the probe behaves correctly.
  - The `_app_lock` serialization (the actually-load-bearing part of 1.7.0's concurrency fix) is unchanged and still active.

## [1.7.0] - 2026-05-18

### Added
- **Concurrent-COM safety in `word_document_server/core/word_com.py`** — a reentrant `_app_lock`, a daemon-thread liveness probe, and two context managers (`undo_record` now wraps the lock+probe; new `com_session` for read-only callers). Two parallel tool calls no longer race on the singleton `_app` (the source of the "transient COM marshalling failures after MCP reconnect" symptom in 1.6.0). When Word is frozen by a modal dialog, calls now raise a descriptive error after `WORD_COM_LIVENESS_TIMEOUT` (default 3 s) instead of hanging the server. Lock acquisition fails fast after `WORD_COM_LOCK_TIMEOUT` (default 30 s); both knobs are env-tunable.
- **Real native endnotes** — `add_endnote_to_document` now writes a proper `<w:endnoteReference>` + `word/endnotes.xml` entry (creating the part, content-type override, relationship, and `EndnoteReference` / `EndnoteText` styles when missing). Previous releases inserted a literal "†" superscript and a fake "Endnotes:" heading paragraph — Word did not recognise those as endnotes. Backed by new `core/footnotes.py::add_endnote_robust`.
- **Real footnote → endnote conversion** — `convert_footnotes_to_endnotes_in_document` now deep-copies each non-separator footnote into `word/endnotes.xml` under a freshly allocated id (skipping pre-existing endnote ids and reserved separator ids -1/0), rewrites every `<w:footnoteReference>` in the body to `<w:endnoteReference>`, and flips the surrounding `rStyle` from `FootnoteReference` to `EndnoteReference`. Previous releases scanned for "¹²³…" superscript runs and appended a fake heading — the output was not a real conversion. Backed by new `core/footnotes.py::convert_footnotes_to_endnotes_robust`.
- **Opt-in persistent osascript session on macOS** (`WORD_MAC_OSA_PERSISTENT=1`) — keeps one `osascript` child alive across JXA calls via a sentinel-framed stdin protocol, dropping per-call overhead from ~100–300 ms to ~5–15 ms. Default remains per-call subprocess for zero regression risk; flip on per-session to validate.

### Fixed
- **`add_footnote_before_text_robust`** — previously delegated to `add_footnote_robust_tool` which has no `position` parameter, so the tool silently inserted the footnote *after* the matched text. Now calls the lower-level `add_footnote_robust` directly with `position="before"`.
- **`add_footnote_to_document`** — removed the legacy fallback path that inserted a literal "¹" character and a fake "Footnotes:" heading when `python-docx.add_footnote()` was unavailable. The fallback produced invalid documents. The tool now always delegates to `add_footnote_robust`.

### Changed
- **`word_document_server/main.py` split** — `register_tools()` (2620 lines of `@mcp.tool` decorators) moved verbatim into a new `word_document_server/tool_registry.py::register_all_tools(mcp)`. `main.py` shrinks from 2808 to 153 lines and now only handles transport configuration, optional `.env` loading, save/path monkeypatches, and server lifecycle. No tool signatures change.
- **`tracked_changes_tools.py` deprecation notice** — the cross-platform OOXML-based `track_replace` / `track_insert` / `track_delete` / `list_tracked_changes` / `accept_tracked_changes` / `reject_tracked_changes` tools now prefix their docstrings with `[DEPRECATED for production: prefer word_live_*]` and document the per-tool migration mapping. Behaviour is unchanged; only the visible description is updated so AI clients route to Word's native revision engine when available.
- **`add_footnote_enhanced` deprecation notice** — same behaviour as `add_footnote_robust`; docstring now points there.
- **`add_footnote_robust_tool` marked CANONICAL** — docstring lists the four older entry points it supersedes.
- **README.md** — top-line marketing now leads with "Native tracked changes / Per-action Ctrl+Z / Data-loss safeguards / Live editing" instead of "124 tools / Cross-platform". The "Data-loss safeguards" section enumerates the control-byte filter, zero-length wildcard guard, 32K InsertBefore auto-chunking, and concurrent-COM lock + liveness probe. The headline tool count is now "100+ tools" with a pointer to `tool_registry.py` for the exact list (reconciling the 76+41+33 / 115 / 124 numbers that previously disagreed across README/TOOLS.md/CHANGELOG).
- **`__version__`** in `word_document_server/__init__.py` synced to `1.7.0` (was stuck at `1.2.0` for several releases).

### Internal
- `XML_NS` constant moved from the bottom of `core/footnotes.py` to the top constants block alongside `W_NS` / `R_NS` / `CT_NS` / `REL_NS`, removing a forward-reference oddity that depended on module-globals late binding.
- Five new helpers in `core/footnotes.py` parallel to the existing footnote helpers: `_get_safe_endnote_id`, `_create_minimal_endnotes_xml`, `_ensure_endnotes_content_type`, `_ensure_endnotes_rels`, `_ensure_endnote_styles`.

## [1.6.0] - 2026-04-29

### Added
- **`word_live_set_core_properties`** — set Word document Title, Subject, Author, Keywords, Comments, Category, Manager, Company, Last Author via `Document.BuiltInDocumentProperties`. Wrapped in `undo_record` so a single Ctrl+Z reverts every property in one call. Equivalent to File > Info > Properties in the Word UI.
- New `word_document_server/utils/text_safety.py` — shared `reject_control_chars()` validator for Find/Replace/Insert text inputs.
- `scrub_orphans` parameter on `word_live_modify_table` `delete_table` operation (default `True`) — cleans stranded `\x07` cell-separator bytes the Word COM `Table.Delete()` occasionally leaves behind.

### Fixed
- **`word_live_replace_text` data-loss vector** — passing `\x07` (cell separator) as `find_text` previously matched across cell boundaries and could delete entire documents. Control bytes (U+0000–U+001F except `\t`, `\n`, `\r`) now rejected with a descriptive error before Find.Execute is reached.
- **`word_live_find_text`** — same control-byte protection applied to `search_text`.
- **`word_live_insert_text`** — same control-byte protection applied to `text` (prevents inserting orphan cell separators outside a real table).
- **`word_live_modify_table` `delete_table`** — leftover `\x07` separators after Word's native `Table.Delete()` now scrubbed by default (configurable via `scrub_orphans=False`).
- **`word_live_add_table`** — rejects `position` offset that falls inside an existing table's range or sits immediately after an orphan cell separator (would otherwise silently merge new content into existing/residual table structure).
- **`word_live_setup_heading_numbering`** — paragraphs that previously kept a custom template style (e.g. `Font Style30/31`) after a forced heading reassignment now (a) get explicit per-paragraph style assignment with try/except, (b) receive the same font/size/bold/color customizations as direct formatting so visual output matches even when the underlying style refuses to swap, (c) report any failed reassignments under a new `restyle_failures` field in the response.
- **`word_live_modify_table`** — re-reads `Tables.Count` per call and validates `table_index` with a helpful message ("table_index N out of range. Document has K table(s)…") instead of throwing "Document has no tables" when a stale index is passed after a prior delete.
- **`word_live_list_open`** — defensive per-document property access; one document in a broken COM proxy state no longer aborts the whole call. Each document entry now includes `index`, `track_revisions`, and per-property `errors` array.
- **`word_live_find_text`** — defensive `Range.Text` / `Range.Start` / `Range.End` / `Document.Name` access via internal `_safe_attr` helper. Transient COM marshalling failures after MCP reconnect now produce partial matches with `<unreadable>` placeholders and a `partial_errors` array, rather than aborting the call.

## [1.5.1] - 2026-04-08

### Fixed
- `word_live_replace_text` — infinite loop when wildcard pattern matches zero-length strings (e.g., `*` alone); now skips forward on zero-length matches and enforces 50K replacement safety ceiling

## [1.5.0] - 2026-04-08

### Added
- **macOS live editing support** via JavaScript for Automation (JXA) — 33 of 41 `word_live_*` tools now work on macOS with Word for Mac
- New module `word_document_server/core/word_mac.py` — JXA bridge with 30+ functions for Word for Mac automation
- Platform auto-detection: same tool names and parameters on both Windows and macOS
- `pywin32` as conditional dependency (Windows only) in `pyproject.toml`

### Changed
- All `print()` calls in `main.py` redirected to stderr — fixes MCP stdio protocol corruption that prevented the server from loading in some clients
- All live tool functions now dispatch to macOS JXA implementations when `sys.platform == "darwin"`
- Updated tool count: 76 cross-platform + 41 Windows Live + 33 macOS Live

### Not Available on macOS
These 4 tools require Windows COM APIs with no AppleScript/JXA equivalent:
- `word_live_get_undo_history` — undo stack inspection not exposed in Word for Mac's scripting dictionary
- `word_live_reply_to_comment` — threaded comment replies not in AppleScript dictionary
- `word_live_resolve_comment` — comment Done property not in AppleScript dictionary
- `word_live_add_watermark` — requires VBA `Shapes.AddTextEffect` (VBA bridge killed by Apple sandboxing in Word 365)

## [1.4.1] - 2026-04-08

### Fixed
- `word_live_replace_text` — `^s` (non-breaking space) now converted to `\u00a0` in replacement text (#4)

## [1.4.0] - 2026-04-08

### Added
- `word_live_insert_paragraphs` — insert multiple paragraphs near a target (by text or index) in a single undo record
- `word_live_take_snapshot` — store paragraph baseline for efficient change detection
- `word_live_get_diff` — compare current document against snapshot, returns only changed paragraphs
- `word_live_snapshot_status` — check snapshot existence and age
- `word_live_modify_table` — new `set_row` and `set_range` operations for bulk cell updates

### Fixed
- `word_live_replace_text` — infinite loop when document has TrackRevisions enabled independently of `track_changes` parameter (#7)
- All destructive tools now unconditionally restore `doc.TrackRevisions` in `finally` block

### Credits
- Snapshot/diff tools, `insert_paragraphs`, and bulk table operations adapted from PR #5 by @FarhadGSRX

## [1.3.0] - 2026-02-28

### Added
- `word_live_modify_table` — table operations via COM: get info, set cell, add/delete rows/columns, merge cells, autofit, delete table
- `word_live_save` — save document in place or save-as (docx, pdf, rtf, txt)
- `word_live_toggle_track_changes` — toggle or explicitly set track changes mode on/off
- `word_live_insert_image` — insert image with sizing, alignment, wrapping, and optional border
- `word_live_insert_cross_reference` — insert live cross-references to headings, bookmarks, figures, tables, equations, footnotes, endnotes
- `word_live_list_cross_reference_items` — list available cross-reference targets with their indices
- `word_live_insert_equation` — insert mathematical equations using UnicodeMath syntax
- `word_live_reply_to_comment` — threaded comment replies (Word 2016+)
- `word_live_resolve_comment` — mark comments as resolved/unresolved (Word 2016+)
- `word_live_delete_comment` — permanently delete a comment
- Total tool count now **114** (75 cross-platform + 39 Windows Live)

### Changed
- `word_live_delete_text` — now table-aware: deletes table objects within range before text deletion
- `word_live_insert_text` — auto-chunks text >30K chars to avoid COM 32K limit
- `word_live_setup_heading_numbering` — handles inflated paragraph ranges from comment anchors
- `word_live_modify_table` set_cell operation now accepts tracked changes before writing to prevent layered content

## [1.2.0] - 2025-02-15

### Added
- `word_live_replace_text` — find & replace via COM that works across tracked change boundaries; supports wildcards (`^m`, `^t`, `^p`) and tracked changes mode
- `word_live_diagnose_layout` — read-only scan for layout problems: keep_with_next chains, heading styles on body text, PageBreakBefore misuse, manual breaks
- `word_live_get_paragraph_format` — inspect paragraph formatting (font, spacing, alignment, list info, style); `include_runs=True` for per-run detail
- `word_live_get_page_text` — read text from specific page(s) with char offsets for chaining into format/edit tools
- `word_live_get_undo_history` — list undo stack entries
- `word_live_apply_list` — apply bullet, numbered, or multilevel list formatting
- `word_live_setup_heading_numbering` — auto-numbered headings (1. / 1.1) via multilevel list linked to Heading styles; configurable style params (font, size, color, spacing)

### Changed
- `word_live_format_text` — added `paragraph_alignment`, `page_break_before`, paragraph-index addressing (`start_paragraph`/`end_paragraph`), `preserve_direct_formatting` for style changes
- `word_live_find_text` — added `use_wildcards` for `^m`/`^t`/`^p`/Word wildcard syntax; `context_chars` now configurable (default 60, was 30)
- `word_live_set_paragraph_spacing` — clarified that `line_spacing` is in points (1.15 lines = 13.8pt)

## [1.1.0] - 2025-01-10

### Added
- 27 Windows Live tools (`word_live_*`) using COM automation for editing documents open in Word
- Per-operation undo system — all destructive tools wrapped with `UndoRecord`; each tool call = one Ctrl+Z entry
- `word_live_undo` — programmatic undo of last N operations
- Live editing tools: `word_live_insert_text`, `word_live_delete_text`, `word_live_format_text`, `word_live_add_table`
- Live reading tools: `word_live_get_text`, `word_live_get_info`, `word_live_find_text`
- Live comment & revision tools: `word_live_add_comment`, `word_live_get_comments`, `word_live_list_revisions`, `word_live_accept_revisions`, `word_live_reject_revisions`
- Live layout tools: `word_live_set_page_layout`, `word_live_add_header_footer`, `word_live_add_page_numbers`, `word_live_add_section_break`, `word_live_set_paragraph_spacing`, `word_live_add_bookmark`, `word_live_add_watermark`
- `word_screen_capture` — screenshot of the Word window
- Cross-platform tracked changes: `track_replace`, `track_insert`, `track_delete`, `list_tracked_changes`, `accept_tracked_changes`, `reject_tracked_changes`
- Cross-platform comments: `add_comment` anchored to text
- Cross-platform hyperlinks: `manage_hyperlinks` (add, list, remove, update)
- Cross-platform layout tools: `set_page_layout`, `add_header_footer`, `add_page_numbers`, `add_section_break`, `set_paragraph_spacing`, `add_bookmark`, `add_watermark`
- Cross-platform footnote tools (10): add, delete, validate, customize footnotes and endnotes
- Cross-platform protection tools: `protect_document`, `unprotect_document`, `add_restricted_editing`, `add_digital_signature`, `verify_document`
- Multiple transport support: stdio (default), SSE, streamable-http
- `MCP_AUTHOR` / `MCP_AUTHOR_INITIALS` environment variables for author metadata
- PyPI packaging as `word-mcp-live`

## [1.0.0] - 2024-12-01

### Added
- Initial release based on [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server)
- 54 cross-platform tools using python-docx
- Document management, content editing, formatting, tables, extraction
- FastMCP server with stdio transport

[1.5.1]: https://github.com/ykarapazar/word-mcp-live/compare/v1.5.0...v1.5.1
[1.5.0]: https://github.com/ykarapazar/word-mcp-live/compare/v1.4.1...v1.5.0
[1.4.1]: https://github.com/ykarapazar/word-mcp-live/compare/v1.4.0...v1.4.1
[1.3.0]: https://github.com/ykarapazar/word-mcp-live/compare/v1.2.0...v1.3.0
[1.2.0]: https://github.com/ykarapazar/word-mcp-live/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/ykarapazar/word-mcp-live/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/ykarapazar/word-mcp-live/releases/tag/v1.0.0
