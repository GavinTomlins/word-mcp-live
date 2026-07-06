# ADR 0006: Template input for Markdown document creation

- **Status:** Accepted
- **Date:** 2026-07-07

## Context

`create_document_from_markdown` (ADR 0004) builds documents with
python-docx's default styles. The first real branded deliverable — an AI
Provenance policy styled with the DArT template — showed the gap: the
output had to inherit *template* styles (fonts, colours, headers/footers,
page setup), so the tool could not be used and the work fell back to a
bespoke script replicating what the external brand-engine skill does
(copy template, clear body, write content with the template's styles).

A second finding from that exercise: real templates are style-sparse. The
DArT template defines Heading 1–8, Title and Normal but not `List Bullet`,
`List Number` or `Table Grid` — the styles the Markdown builder assumed.

## Decision

Add `template=` to `markdown_to_document` and the
`create_document_from_markdown` tool:

1. The output starts as a byte copy of the template, so styles, headers,
   footers, numbering definitions and page setup are inherited exactly.
2. The template's body content is cleared (everything except the trailing
   `sectPr`), making the template a *format* source, never a *content*
   source.
3. Style lookups degrade gracefully when the template does not define
   them: list styles fall back through `List Bullet/Number [n]` →
   `List Paragraph` → indented literal markers; tables fall back from
   `Table Grid` to directly-applied cell borders; missing heading styles
   fall back to bold paragraphs. The document must render correctly in
   every template, not just rich ones.

## Process

Driven by the AI Provenance build (2026-07-07): the bespoke script's
template-copy approach was promoted into the core module, and its
border/bullet workarounds became the fallback tier. Verified by unit tests
(style inheritance, style-sparse fallbacks, missing-template rejection)
and an end-to-end run against the real DArT template (Atkinson
Hyperlegible headings, brand footer, clean validation).

## Consequences

- Branded deliverables are one tool call: Markdown + template path.
- The template's body is never merged — callers wanting template text
  must put it in the Markdown. This keeps the "template = form" rule
  unambiguous.
- Fallback rendering means output fidelity varies with template richness;
  the validation/read-back loop remains the arbiter of "done".
