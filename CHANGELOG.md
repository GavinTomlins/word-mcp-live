# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- **`validate_document`** — business-rule validation of the OOXML package with actionable findings: table `tcW` vs `gridCol` width mismatch ("will skew"), image display vs actual aspect ratio ("distorted"), comment anchor/definition integrity across the comment part files, TOC fields without `w:updateFields`, missing `xml:space="preserve"`, and `pPr`/`rPr` element-order violations (ADR 0003).
- **`get_document_markdown`** — structural read-back of a document as Markdown (headings, lists, tables, emphasis, hyperlinks); `show_revisions=True` renders tracked changes as CriticMarkup `{++ins++}`/`{--del--}` (ADR 0004).
- **`create_document_from_markdown`** — batch document creation from a Markdown subset (headings, emphasis, code, links, nested bullet/numbered lists, pipe tables, `---` rendered as a border paragraph) in a single call (ADR 0004). Accepts `template=` to inherit an existing .docx's styles, headers/footers and page setup — with graceful fallbacks for style-sparse templates (ADR 0006).
- **`set_update_fields_on_open`** — sets `w:updateFields` in settings.xml so Word refreshes TOC/page-number fields on open (ADR 0004).
- **`word_workflow_guide` MCP prompt** — tool-routing and verification workflow served over the protocol to any client (ADR 0005).
- **Project skill and agent** — `.claude/skills/word-live/SKILL.md` (routing decision tree + guardrails) and `.claude/agents/doc-oracle.md` (Word-document specialist with a mandatory validate/read-back loop) ship with the repo (ADR 0005).
- **Architecture decision records** — `docs/adr/` documents the FastMCP 3.x architecture and the design lessons behind the new quality tools.

### Fixed
- **Startup crash on FastMCP 3.x** — `FastMCP(description=...)` raised `TypeError` (the parameter is `instructions`); the server could not start with the locked `fastmcp` 3.4.3. Also fixed double-escaped `\n` in the server instructions.
- **Startup crash on macOS/Linux** — `screen_capture_tools` imported PIL at module level while Pillow is a Windows-only dependency; the import is now deferred to the Windows capture path.
- **All `word_live_*` tools broken on macOS** — a misapplied patch left `word_mac.py` uncompilable (Python statements pasted inside a VBA-lines list literal, dropping the `fileNum` declaration); every live tool failed at import with "invalid syntax (line 1311)".
- **`word_live_save` reported success without saving on macOS** — Word for Mac's scripted save can silently no-op when the sandbox wants user consent, while the document's `saved` property still reads true. `mac_save` now verifies the on-disk mtime moved (or the saveAs target exists) and returns `saved: false` with recovery guidance when the write never landed.
- **`word_live_format_text` highlight crash on macOS** — Windows `WdColorIndex` integers were fed to the JXA string escaper; they are now mapped to the enum name strings JXA expects (`7` → `"yellow"`, etc.).

### Changed
- **FastMCP 3.x alignment** (per [gofastmcp.com/servers/server](https://gofastmcp.com/servers/server)): pinned `fastmcp>=3.0,<4`; server is now built by a `build_server(settings)` factory (`server.py`) with `instructions`, `on_duplicate="error"`, and `mask_error_details` (on by default for HTTP).
- **Native authentication** — the hand-rolled ASGI bearer middleware was replaced by FastMCP's `StaticTokenVerifier` via the `auth=` constructor parameter (standard 401 + `WWW-Authenticate` handling).
- **Typed configuration** — new `config.py` (pydantic-settings). New `WORD_MCP_*` variable names with full backward compatibility for the legacy names (`MCP_TRANSPORT`, `MCP_HOST`, `PORT`/`MCP_PORT`, `MCP_PATH`, `WORD_MCP_LIVE_API_KEY`, `WORD_MCP_LIVE_INSECURE`, `MCP_AUTHOR`, `MCP_AUTHOR_INITIALS`).
- **Modular tool registration** — the 1,900-line `register_tools()` in `main.py` was split into 13 modules under `mcp_tools/`, each exposing `register(mcp)`.
- **Platform-aware live tools** — all live tools are tagged `live` and are hidden (not just failing) on platforms without a Word COM/JXA bridge, via `mcp.disable(tags={"live"})`.
- **Lifespan management** — save/path monkey-patches and the automatic backup loop now run in the FastMCP lifespan (with teardown) instead of imperatively in `run_server()`.
- **Observability** — new per-tool logging middleware (name, latency, outcome) and an unauthenticated `GET /health` endpoint for platform health checks.

- **Packaging** — `pytest` moved from runtime dependencies to a new `dev` extra, so end users no longer install the test framework and CONTRIBUTING's `pip install -e ".[dev]"` works as documented.
- **Fork rebrand** — distribution renamed to `word-mcp-live-gavintomlins` and the internal package to `word_mcp_live_gavintomlins`; registry metadata (`server.json`, `manifest.json`), client config examples, and contact links updated to this fork, with the full lineage credited in README Acknowledgments and the LICENSE copyright holders retained (Gavin Tomlins added). The `word_mcp_server` entry point is unchanged, so existing client configs keep working.

### Removed
- **SSE transport** (deprecated in the MCP spec) — use `WORD_MCP_TRANSPORT=http`. `MCP_TRANSPORT=sse` now fails fast with a clear error; `setup_mcp.py` no longer offers SSE.

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

[1.5.1]: https://github.com/cheemscheems/word-mcp-live/compare/v1.5.0...v1.5.1
[1.5.0]: https://github.com/cheemscheems/word-mcp-live/compare/v1.4.1...v1.5.0
[1.4.1]: https://github.com/cheemscheems/word-mcp-live/compare/v1.4.0...v1.4.1
[1.3.0]: https://github.com/cheemscheems/word-mcp-live/compare/v1.2.0...v1.3.0
[1.2.0]: https://github.com/cheemscheems/word-mcp-live/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/cheemscheems/word-mcp-live/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/cheemscheems/word-mcp-live/releases/tag/v1.0.0
