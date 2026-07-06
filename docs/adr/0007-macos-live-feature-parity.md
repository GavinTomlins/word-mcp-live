# ADR 0007: Closing the macOS feature gaps (threading, resolve, watermark)

- **Status:** Accepted
- **Date:** 2026-07-07

## Context

Four features were Windows-only: threaded comment replies, comment
resolve/unresolve, undo-history inspection, and watermarks. Probing Word
for Mac 16.107's actual scripting surface established why, and killed two
tempting approaches:

- The AppleScript/JXA dictionary's `Word comment` class exposes neither
  `done` nor `replies`; there is no undo-stack API and no WordArt command.
- **`do Visual Basic` has been removed** from the dictionary (only
  `run VB macro`, which invokes pre-existing named macros, remains) — so
  the classic inline-VBA escape hatch is gone, and the one legacy
  `do Visual Basic` call in the codebase is dead at runtime.
- **.NET is not an alternative for live editing**: COM does not exist on
  macOS, so no .NET runtime can attach to a running Word process. Its only
  unique value here would be the OpenXML validator (deferred in ADR 0003).
- The scripted `save` and `close` verbs are also degraded in 16.107,
  reinforcing that live-layer features must minimise dictionary surface.

## Decision

1. **Comment threading and resolution go file-based** (closed document),
   implemented over the full multi-part model in
   `core/comment_threads.py`: `comments.xml` (+ generated `w14:paraId`),
   `commentsExtended.xml` (`w15:done`, `w15:paraIdParent`),
   `commentsIds.xml` and `commentsExtensible.xml` (durable ids), with
   reply anchors co-located with the parent's range. New tools:
   `reply_to_comment`, `resolve_comment`, `delete_comment` (cascading to
   replies). The comment readers were rebuilt on the same package-level
   parser — the legacy python-docx reader could not see comments at all
   (generic part without `.element`; fallback string-tested an lxml repr)
   — and now report `resolved`, `parent_id` and `reply_count`.
2. **Live watermarks work on macOS via the dictionary itself**: a shape
   created *at the header-footer object* anchors in the primary header
   story (repeats on every page) — a transparent, borderless rectangle,
   rotated, with large gray text and behind-text wrap, centred from page
   geometry (`mac_add_watermark`). The COM magic-centring constants are
   rejected on Mac, so position is computed from `page setup`.
3. **Undo-history inspection stays Windows-only** for now; a server-side
   operation journal is the recorded follow-up, out of scope here.
4. `.dotm` + `run VB macro` and an Office.js sidecar add-in remain the
   documented escalation paths if live threading/resolve on macOS is ever
   required.

## Process

Every claim above was probed empirically against the running Word 16.107
(sdef inspection plus live scripting on scratch documents): the header
anchoring was discovered by testing `make new shape at <header>` after
selection-based anchoring landed in the main story; property names and
enum spellings (`autoshape rectangle`, `wrap behind`, 0–65535 AppleScript
RGB) were validated the same way. The file-based threading was verified by
unit tests and by threading/resolving a real reviewed document, with the
business-rule validator confirming part integrity.

## Consequences

- macOS (and Linux) gain replies/resolve/delete with the closed-file
  caveat; macOS gains live watermarks.
- Comment listings are now thread-aware everywhere.
- Live comments created through Word carry the signed-in account as
  author (Word's behaviour); file-based comments honour `WORD_MCP_AUTHOR`.
