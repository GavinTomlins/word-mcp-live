# ADR 0004: Markdown as the verification and batch-authoring format

- **Status:** Accepted
- **Date:** 2026-07-06

## Context

Two gaps were identified by comparing the server against Kimi's docx skill
and a second widely deployed assistant docx skill:

1. **Verification.** Kimi's rule is "pandoc is the SOURCE OF TRUTH — the
   validator checks structure; pandoc shows actual content", and it requires
   a pandoc round-trip before any delivery. The server's closest facility was
   `get_document_text`, which loses headings, lists, and tables — an agent
   cannot confirm *structure* survived, only words.
2. **Creation.** Both skills build new documents as one program (C#/OpenXML
   or docx-js) rather than incremental calls. On this server a modest report
   takes ~30 `add_paragraph`/`add_heading`/`add_table` round-trips — slow,
   token-expensive, and a failure midway leaves a half-built file.

## Decision

Use Markdown as the common structured format in both directions, with no new
dependencies (a python-docx walker/builder rather than pandoc):

- **`get_document_markdown(filename, show_revisions=False)`** renders
  headings, lists, tables, bold/italic and hyperlinks; with
  `show_revisions=True`, tracked insertions render as `{++text++}` and
  deletions as `{--text--}` (CriticMarkup, matching pandoc's
  `--track-changes=all` idiom the skills verify against).
- **`create_document_from_markdown(filename, markdown, ...)`** builds a
  complete document in one call from a Markdown subset: `#`–`######`
  headings, paragraphs, `**bold**`/`*italic*`/`` `code` ``, nested bullet and
  numbered lists, pipe tables (bold header row), links, and `---` rendered as
  a bottom-border paragraph — never a table used as a divider (a documented
  docx-js pitfall in the second skill).
- **`set_update_fields_on_open(filename)`** sets `w:updateFields` in
  settings.xml so TOC/page fields refresh on open — Kimi calls the
  equivalent (`SetUpdateFieldsOnOpen`) unconditionally in every generated
  document.

Markdown is the interchange format because it is what models author most
reliably, diff cleanly, and already use for the revision-markup convention.

## Process

Derived from the skills analysis: Kimi's mandatory pandoc verification loop
and single-program creation pipeline; the second skill's docx-js batch
creation and its divider/table pitfall list. Fidelity target was set at "an
agent can confirm every structural element it authored", not "lossless OOXML
round-trip".

## Consequences

- The create → validate → read-back loop is three tool calls total.
- The Markdown subset is intentionally partial (no images, footnotes or
  nested tables in v1); the incremental tools remain the escape hatch.
- If pandoc is present on a host nothing changes — these tools remove the
  dependency, not the option.
