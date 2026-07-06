"""Workflow guidance served over MCP as a prompt (ADR 0005)."""

WORKFLOW_GUIDE = """\
# Word Document Server — workflow guide

## Choosing tools

1. **File CLOSED in Word** → use the cross-platform tools (no prefix):
   create_document_from_markdown, add_paragraph, search_and_replace,
   track_replace, add_comment, get_document_markdown, validate_document, …
2. **File OPEN in Word** (or a cross-platform tool failed with a file-lock
   error) → switch to the corresponding `word_live_*` tool. Every live
   action is one Ctrl+Z undo step for the user.
3. `word_live_*` tools exist only on Windows and macOS; on other platforms
   they are hidden.

## Creating a document

1. Prefer **one** `create_document_from_markdown` call over many
   incremental add_* calls.
2. For branded output, pass `template=<path.docx>` — the new document
   inherits the template's styles, headers/footers and page setup (the
   template's body content is not carried over).
3. If a template is provided, act as a form-filler, not a designer: replace
   placeholders, keep the format.
4. Added a table of contents or page-number fields? Call
   `set_update_fields_on_open` so readers do not see stale values.

## Verifying — work is not done until verified

1. `validate_document` → fix every error; treat warnings seriously
   (skewed tables, distorted images, invisible comments).
2. `get_document_markdown` → confirm the structure and text match intent.
   With `show_revisions=True`, insertions show as {++text++} and deletions
   as {--text--}.
3. For live sessions: `word_live_take_snapshot` before a change,
   `word_live_get_diff` after — report the diff, not just "done".

## Editing etiquette

- `backup_document` before any large or destructive change.
- Never edit inside another author's tracked insertion or deletion:
  reject their insertion by nesting a deletion; restore their deletion by
  inserting after it.
- Keep the document language identical to the conversation language —
  including headings, headers/footers, TOC text and labels.
"""


def register(mcp):
    """Register guidance prompts on ``mcp``."""

    @mcp.prompt
    def word_workflow_guide() -> str:
        """Decision guide for creating, editing and verifying Word documents
        with this server: closed-file vs live tools, batch creation, and the
        mandatory validate + read-back loop."""
        return WORKFLOW_GUIDE
