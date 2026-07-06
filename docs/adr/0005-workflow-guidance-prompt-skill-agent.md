# ADR 0005: Ship workflow guidance with the server (prompt, skill, agent)

- **Status:** Accepted
- **Date:** 2026-07-06

## Context

The server's hardest-won knowledge is procedural, not mechanical: when to use
cross-platform vs `word_live_*` tools, what to do on a file-lock error, that
work is not done until it has been validated and read back. Today that
knowledge lives only in this repository's project-instructions file, so it
reaches only interactive sessions opened in this repository. An audit of the
consuming environment found the associated agent repertoire had **zero**
integration with this server — no agent, skill, or MCP config referenced it —
while two file-based docx skills solved overlapping problems in parallel.

Kimi's skill demonstrates how much output quality comes from prompt-level
guardrails rather than code: "template provided = act as form-filler, not
designer"; document language must match conversation language everywhere;
word counts within ±20% without bullet-point padding; scene completeness
(contracts get signature blocks, minutes get action items with owners).

## Decision

Ship the guidance at three layers, each reaching a different consumer:

1. **MCP prompt** (`mcp_tools/guidance.py`) — a `word_workflow_guide` prompt
   served over the protocol itself, so *any* MCP client can pull the
   decision tree: closed file → cross-platform tools; open file / lock error
   → `word_live_*`; always backup before destructive edits; always
   validate + read back before reporting success.
2. **Project skill** (`.claude/skills/word-live/SKILL.md`) — the routing
   decision tree plus the Kimi-derived guardrails, for interactive sessions
   in this repository.
3. **Project agent** (`.claude/agents/doc-oracle.md`) — a Word-document
   specialist whose system prompt hard-codes the verification loop and
   tracked-changes etiquette (never edit inside another author's revision;
   use reject/restore semantics).

## Process

The three-layer split follows from the audit: the MCP prompt covers remote
clients and any MCP host, the skill covers interactive sessions, the agent
covers delegated work. Guardrail content was selected from the Kimi skill's
SKILL.md rules that are behavioural (promptable) rather than mechanical
(already implemented as tools), and from the second skill's redlining
etiquette.

## Consequences

- Any MCP client can discover the intended workflow without repo access.
- The skill/agent files version with the server, so guidance and tool
  surface cannot drift apart silently.
- Wiring the server into other agent runtimes (e.g. an opencode
  `mcp_config.json`) remains a per-environment step, deliberately out of
  scope here.
