# ADR 0001: Record architecture decisions

- **Status:** Accepted
- **Date:** 2026-07-06

## Context

The server has been through one major restructure (FastMCP 3.x alignment) and
is now absorbing design lessons from external document-generation systems
(Kimi's docx skill and a second widely deployed assistant docx skill).
Decisions made during these
efforts live in commit messages and conversation history, which do not
survive as a browsable engineering record.

## Decision

Record significant architecture decisions as ADRs in `docs/adr/`, numbered
sequentially, using this lightweight format: Status, Date, Context, Decision,
Process (how the decision was reached, including sources analysed), and
Consequences.

## Consequences

- New contributors can trace *why* the server is shaped the way it is.
- Each substantial change series should land with at least one ADR.
- ADRs are immutable once accepted; superseding decisions get a new ADR that
  references the old one.
