# ADR 0003: Two-layer document validation as a first-class tool

- **Status:** Accepted
- **Date:** 2026-07-06

## Context

The server exposed ~120 tools that write documents but none that check them.
A tool call returning "success" only means the write completed — not that the
document opens in Word, that tables render at the intended widths, or that
comments are actually visible. Two mature document-generation systems were
analysed for comparison:

- **Kimi's docx skill** runs a mandatory two-layer pipeline after every
  build: an OpenXML schema validator, then custom business rules — table
  `gridCol` width vs cell `tcW` mismatch ("will skew"), image display vs
  actual aspect ratio ("distorted"), comment multi-file sync, element-order
  violations ("missing tblGrid → Word cannot open").
- **A second widely deployed assistant docx skill** validates on every
  repack, auto-repairs
  mechanical faults, and — critically — compares against the original file so
  that editing an already-imperfect document only fails on *new* errors.

## Decision

Add a `validate_document` tool backed by a new `core/validation.py` module
implementing the business-rule layer over the OOXML package directly (lxml,
no new dependencies):

- **Package integrity** — the file unzips and `document.xml` parses.
- **Tables** — `tblGrid` present; grid column count vs row cell count
  (gridSpan-aware); `tcW` vs `gridCol` width mismatch beyond ±5% → "will
  skew" warning.
- **Images** — inline drawing display ratio vs actual pixel ratio of the
  embedded image (PNG/JPEG/GIF headers parsed directly) beyond ±5% →
  "distorted" warning.
- **Comments** — anchors in `document.xml` vs definitions in `comments.xml`
  in both directions; threaded replies without `commentsExtended.xml`.
- **Fields** — a TOC field without `w:updateFields` in settings → warning
  pointing at `set_update_fields_on_open`.
- **Text** — `w:t` with leading/trailing whitespace missing
  `xml:space="preserve"`.
- **Element order** — children of `pPr`/`rPr` out of OOXML schema order.

Findings are returned as structured errors/warnings with Kimi-style
actionable messages (location + measured values + consequence), so the model
reading the report can fix the document without further diagnosis.

Full XSD schema validation is deliberately out of scope: it needs the .NET
OpenXML SDK or bundled schemas, and the live-editing path already gets
schema-equivalent validation for free (Word itself is holding the document).

## Process

Both skills' validator sources were read end-to-end
(`validate_docx.py`/`business_rules.py`/`fix_element_order.py` from Kimi;
`validators/docx.py`/`base.py`/`redlining.py` from the second skill). The
check list above is the intersection of what both systems treat as mandatory
plus the highest-value items unique to each, filtered to what is
implementable with the server's existing dependency set.

## Consequences

- Agents can (and per the bundled skill, must) verify documents before
  declaring work done.
- Validation is advisory, not blocking — write tools stay fast; callers
  choose when to pay the validation cost.
- The ±5% thresholds match Kimi's field-tested values and are encoded as
  constants for tuning.
