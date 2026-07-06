---
name: word-live
description: Route Word document work to the right word-mcp-live tools (cross-platform vs live), and enforce the create → validate → read-back verification loop. Use whenever creating, editing, reviewing, or reading .docx files with this server's MCP tools.
---

# word-live — working with Word documents through word-mcp-live

## Routing decision tree

1. **Is the document open in Word right now** (user is watching, or a
   cross-platform tool just failed with a file-lock error)?
   → Use `word_live_*` tools. Every action is a single Ctrl+Z for the user.
2. **Is the file closed** (batch creation, unattended processing)?
   → Use the cross-platform tools (no prefix).
3. **Creating a new document?**
   → One `create_document_from_markdown` call, not a chain of
   `add_paragraph`/`add_heading` calls. Fall back to incremental tools only
   for things Markdown can't express (images, footnotes, protection).
4. **Need raw OOXML surgery beyond any tool?**
   → `get_document_xml` to inspect; escalate to an unpack/edit/pack pipeline
   only as a last resort.

## The verification loop — work is not done until verified

After creating or substantially editing a document:

1. `validate_document` — fix all errors; take warnings seriously (skewed
   table widths, distorted images, invisible comments, stale TOC fields).
2. `get_document_markdown` — read the document back and compare against
   intent: headings present? tables filled? emphasis where expected?
   Use `show_revisions=True` to see `{++insertions++}` / `{--deletions--}`.
3. Live sessions: `word_live_take_snapshot` before the change,
   `word_live_get_diff` after — report what actually changed.

Never report success from write-tool return values alone.

## Guardrails

- **Backup first**: `backup_document` before encryption, large replaces, or
  any destructive operation.
- **Template = form-filler**: when the user supplies a template, replace
  placeholders only; the format is their decision. No template = you are the
  designer.
- **Fields refresh**: after `add_table_of_contents` or page-number fields,
  call `set_update_fields_on_open`.
- **Language consistency**: document language matches the conversation
  language everywhere — headings, headers/footers, TOC text, table labels.
- **Tracked-changes etiquette**: never edit inside another author's
  insertion or deletion. Reject their insertion by nesting a deletion;
  restore their deletion by inserting after it. Attribute your changes with
  the configured author (`WORD_MCP_AUTHOR`).
- **Length discipline**: hit requested word counts within ±20%; never pad
  with bullet lists.
- **Scene completeness**: think ahead about what the document type needs —
  contracts get signature/date blocks for both parties; minutes get
  attendees and action items with owners; exams get name/ID fields and
  point allocations.

## Common failure → recovery

| Symptom | Recovery |
| --- | --- |
| `File locked (probably open in Word)` | Switch to the matching `word_live_*` tool |
| `word_live_*` tool missing from tool list | Platform has no Word bridge (Linux) — ask the user to close the file and use cross-platform tools |
| Validation reports `will skew` / `distorted` | Fix widths with `set_table_column_widths` / re-insert the image at its native aspect ratio |
| TOC shows wrong page numbers | `set_update_fields_on_open`, then reopen in Word |
