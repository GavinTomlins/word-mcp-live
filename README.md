<div align="center">

# word-mcp-live-gavintomlins

**The only MCP server that edits Word documents while they're open**

`Live editing` &middot; `Tracked changes` &middot; `Per-action undo` &middot; `125 tools` &middot; `Cross-platform`

[![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Platform: Windows + macOS/Linux](https://img.shields.io/badge/platform-Windows%20%2B%20macOS%2FLinux-lightgrey)]()

</div>

---

word-mcp-live-gavintomlins gives any AI assistant that supports [MCP](https://modelcontextprotocol.io/) full control of Microsoft Word. Open a document, tell the AI what you need, and watch it happen — formatting, tracked changes, comments, and all. Changes appear live in your open document.

## What's New — GT Improvements

The server core was rebuilt on FastMCP 3.x (full details in the [CHANGELOG](CHANGELOG.md)):

- **FastMCP 3.x** — pinned `fastmcp>=3.0,<4`; the server is built by a `build_server(settings)` factory with lifespan-managed document hooks and backups, plus per-tool logging middleware.
- **Typed configuration** — new `WORD_MCP_*` environment variables (pydantic-settings); all legacy names (`MCP_TRANSPORT`, `WORD_MCP_LIVE_API_KEY`, `MCP_AUTHOR`, …) keep working as aliases, so existing configs need no changes.
- **Streamable HTTP** — `WORD_MCP_TRANSPORT=http` serves the modern streamable HTTP transport with native bearer-token authentication and an unauthenticated `GET /health` endpoint. The deprecated SSE transport was removed.
- **Platform-aware tools** — live editing tools are tagged `live` and automatically hidden on platforms without a Word COM/JXA bridge (Linux), instead of failing at call time.
- **Modular registration** — tool registrations live in `word_mcp_live_gavintomlins/mcp_tools/`, one module per category.
- **Error masking** — exception details are masked from clients by default on HTTP deployments.
- **Document validation** — `validate_document` runs business-rule checks (skewed table widths, distorted images, comment integrity, stale TOC fields, OOXML element order) with actionable messages ([ADR 0003](docs/adr/0003-two-layer-document-validation.md)).
- **Markdown in and out** — `create_document_from_markdown` builds a full document in one call; `get_document_markdown` reads structure back for verification, with `{++ins++}`/`{--del--}` revision markup; `set_update_fields_on_open` keeps TOC/page fields fresh ([ADR 0004](docs/adr/0004-markdown-verification-and-authoring.md)).
- **Workflow guidance shipped with the server** — a `word_workflow_guide` MCP prompt, a `word-live` project skill, and the **doc-oracle agent** (see below) encode tool routing and the mandatory validate + read-back loop ([ADR 0005](docs/adr/0005-workflow-guidance-prompt-skill-agent.md)).
- **Architecture decision records** — design decisions and the processes behind them are documented in [docs/adr/](docs/adr/).

<table>
<tr>
<td width="50%">

### Without word-mcp-live-gavintomlins

- AI can discuss your document but can't touch it
- You copy-paste between AI and Word, losing formatting
- Track changes? You do those manually after the fact
- Every edit means save → close → process → reopen

</td>
<td width="50%">

### With word-mcp-live-gavintomlins

- "Add a tracked change replacing ABC Corp with XYZ Ltd" — done
- Changes appear live in your open Word document
- Every AI edit is one Ctrl+Z away
- Real tracked changes with your name, not XML hacks

</td>
</tr>
</table>

## What Sets This Apart

- **Live editing** — Edit documents while they're open in Word. No save-close-reopen cycle.
- **Full undo** — Every AI action is a single Ctrl+Z. Made a mistake? Just undo it.
- **Native tracked changes** — Real Word revisions with your name, not XML hacks.
- **Threaded comments** — Add, reply, resolve, and delete comments like a human reviewer.
- **Layout diagnostics** — Detects formatting problems before they become print disasters.
- **Equations & cross-references** — Insert math formulas and auto-updating references.
- **125 tools** — The most comprehensive Word MCP server available.
- **Automatic backups** — Periodic backup every 5 minutes + on-demand backup before destructive operations. Stored in `_backup/` folder, max 5 copies kept.
- **Path sandbox** — Optional `MCP_ALLOWED_DIR` restricts file access to a single directory tree.
- **COM timeout protection** — Long-running COM operations (replace, save, etc.) have configurable timeouts to prevent server hang on remote/unstable Word connections.
- **Security-hardened** — All 10 audit findings from the initial security review have been addressed (fake signatures removed, AppleScript injection fixed, path traversal prevented, predictable temp files eliminated, and more).

## Quick Start

```bash
pip install git+https://github.com/GavinTomlins/word-mcp-live.git
```

Or install from source:

```bash
git clone https://github.com/GavinTomlins/word-mcp-live.git
cd word-mcp-live
pip install -e .
```

## Client Installation

<details open>
<summary><b>Claude Desktop</b></summary>

Add to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "word-mcp-live-gavintomlins": {
      "command": "uvx",
      "args": [
        "--from",
        "git+https://github.com/GavinTomlins/word-mcp-live.git",
        "word_mcp_server"
      ],
      "env": {
        "WORD_MCP_AUTHOR": "Your Name",
        "WORD_MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

</details>

<details>
<summary><b>Claude Code</b></summary>

Add to your `.mcp.json`:

```json
{
  "mcpServers": {
    "word-mcp-live-gavintomlins": {
      "command": "uvx",
      "args": [
        "--from",
        "git+https://github.com/GavinTomlins/word-mcp-live.git",
        "word_mcp_server"
      ],
      "env": {
        "WORD_MCP_AUTHOR": "Your Name",
        "WORD_MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

</details>

<details>
<summary><b>Cursor</b></summary>

Add to `~/.cursor/mcp.json`:

```json
{
  "mcpServers": {
    "word-mcp-live-gavintomlins": {
      "command": "uvx",
      "args": [
        "--from",
        "git+https://github.com/GavinTomlins/word-mcp-live.git",
        "word_mcp_server"
      ],
      "env": {
        "WORD_MCP_AUTHOR": "Your Name",
        "WORD_MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

</details>

<details>
<summary><b>VS Code / Copilot</b></summary>

Add to your VS Code `settings.json`:

```json
{
  "mcp": {
    "servers": {
      "word-mcp-live-gavintomlins": {
        "command": "uvx",
        "args": [
          "--from",
          "git+https://github.com/GavinTomlins/word-mcp-live.git",
          "word_mcp_server"
        ],
        "env": {
          "WORD_MCP_AUTHOR": "Your Name",
          "WORD_MCP_AUTHOR_INITIALS": "YN"
        }
      }
    }
  }
}
```

</details>

<details>
<summary><b>Windsurf</b></summary>

Add to `~/.codeium/windsurf/mcp_config.json`:

```json
{
  "mcpServers": {
    "word-mcp-live-gavintomlins": {
      "command": "uvx",
      "args": [
        "--from",
        "git+https://github.com/GavinTomlins/word-mcp-live.git",
        "word_mcp_server"
      ],
      "env": {
        "WORD_MCP_AUTHOR": "Your Name",
        "WORD_MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

</details>

<details>
<summary><b>Docker</b></summary>

> An image is not yet published for this fork. Once available it will be:

```json
{
  "mcpServers": {
    "word-mcp-live-gavintomlins": {
      "command": "docker",
      "args": ["run", "-i", "--rm", "ghcr.io/gavintomlins/word-mcp-live"],
      "env": {
        "WORD_MCP_AUTHOR": "Your Name",
        "WORD_MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

> Note: a container only exposes the cross-platform tools. Live editing needs Word on the host (Windows COM / macOS JXA), which a container cannot reach.

</details>

> **`WORD_MCP_AUTHOR`** sets your name on tracked changes and comments (default: `"Author"`; legacy alias `MCP_AUTHOR`). **`WORD_MCP_AUTHOR_INITIALS`** sets comment initials.

## Two Modes

|  | Works everywhere | Live editing (Word open) |
|---|---|---|
| **What it does** | Create and edit saved .docx files | Edit documents live while you work in Word |
| **Platform** | Windows, macOS, Linux | Windows (COM) and macOS (JXA) |
| **Undo** | File-level saves | Per-action Ctrl+Z (Windows); per-operation undo (macOS) |
| **Best for** | Batch processing, document generation | Interactive editing, formatting, review |

Both modes work together. The AI picks the right one for the task.

## The doc-oracle Agent

This repo ships a ready-to-use Word-document specialist for agent runtimes:

- **[doc-oracle](.claude/agents/doc-oracle.md)** — an agent whose system prompt hard-codes the working discipline: route between cross-platform and live tools, build new documents with a single `create_document_from_markdown` call, back up before destructive edits, respect other authors' tracked changes, and **never report success without running `validate_document` and reading the result back** with `get_document_markdown` (or `word_live_get_diff` in live sessions).
- **[word-live skill](.claude/skills/word-live/SKILL.md)** — the same routing decision tree and guardrails for interactive sessions.
- **`word_workflow_guide` MCP prompt** — served over the protocol itself, so any MCP client can pull the workflow without repo access.

Delegate document work to doc-oracle when you want the verification loop enforced rather than remembered.

### macOS Live Editing (New in v1.5.0)

Live tools now work on macOS via JavaScript for Automation (JXA). Same tool names, same parameters — the server detects your platform and uses the right automation backend.

| Feature | Windows | macOS |
|---------|---------|-------|
| Text read/write/find/replace | COM | JXA |
| Formatting (bold, font, style) | COM | JXA |
| Track changes & revisions | COM | JXA |
| Comments (add, delete, list) | COM | JXA |
| Tables (read, write, add rows) | COM | JXA |
| Page layout, headers, bookmarks | COM | JXA |
| Equations, cross-references | COM | JXA |
| Threaded comment replies | COM | Not available |
| Comment resolve/unresolve | COM | Not available |
| Undo history inspection | COM | Not available |
| Watermarks | COM | Not available |

## Configuration

| Variable | Legacy alias | Default | Description |
|----------|--------------|---------|-------------|
| `WORD_MCP_AUTHOR` | `MCP_AUTHOR` | `"Author"` | Author name for tracked changes and comments |
| `WORD_MCP_AUTHOR_INITIALS` | `MCP_AUTHOR_INITIALS` | `""` | Author initials for comments |
| `WORD_MCP_TRANSPORT` | `MCP_TRANSPORT` | `stdio` | Transport type: `stdio` or `http` (streamable HTTP; `streamable-http` also accepted). SSE was removed (deprecated in the MCP spec) |
| `WORD_MCP_HOST` | `MCP_HOST` | `127.0.0.1` | Host to bind (HTTP transport; use `0.0.0.0` for remote access) |
| `WORD_MCP_PORT` | `PORT`, `MCP_PORT` | `8000` | Port to bind (HTTP transport) |
| `WORD_MCP_PATH` | `MCP_PATH` | `/mcp` | Endpoint path (HTTP transport) |
| `WORD_MCP_API_KEY` | `WORD_MCP_LIVE_API_KEY` | *(required for HTTP)* | Bearer token for HTTP transport authentication. Set to a secret value |
| `WORD_MCP_INSECURE` | `WORD_MCP_LIVE_INSECURE` | `false` | Set to `true` to disable authentication (local/dev only, NOT for remote access) |
| `WORD_MCP_LOG_LEVEL` | `FASTMCP_LOG_LEVEL` | `INFO` | Log level (`DEBUG`, `INFO`, `WARNING`, `ERROR`) |
| `WORD_MCP_MASK_ERRORS` | — | *(auto)* | Mask exception details from clients; defaults to on for HTTP, off for stdio |
| `MCP_ALLOWED_DIR` | — | *(none)* | Restrict file access to this directory and its subdirectories (path sandbox) |
| `MCP_MAX_BACKUPS` | — | `5` | Max automatic backups to keep per document; set to `0` for unlimited |

The HTTP transport also exposes an unauthenticated `GET /health` endpoint for
load-balancer and platform health checks.

For remote deployment, see [RENDER_DEPLOYMENT.md](RENDER_DEPLOYMENT.md).

## Example Prompts

Just tell the AI what you want in plain language:

```
"Draft a contract with tracked changes so my colleague can review"
"Format all headings as Cambria 13pt bold and add automatic numbering"
"Add a comment on paragraph 3 asking about the deadline"
"Find every mention of 'ABC Corp' and replace with 'XYZ Ltd' as a tracked change"
"Set the page to A4 landscape with 2cm margins"
"Insert a table of contents based on the document headings"
"Add page numbers in the footer and our company name in the header"
"Insert a cross-reference to Heading 2 in paragraph 5"
```

## Usage Examples

### Example 1: Read a document (cross-platform)

**Tool call:** `get_document_text`
```json
{ "filename": "C:/Documents/report.docx" }
```
**Expected output:**
```json
{
  "status": "success",
  "paragraphs": [
    {"index": 0, "text": "Quarterly Report", "style": "Heading 1"},
    {"index": 1, "text": "Revenue increased by 15% compared to Q3.", "style": "Normal"},
    {"index": 2, "text": "Key Metrics", "style": "Heading 2"}
  ],
  "total_paragraphs": 3
}
```

### Example 2: Live editing with tracked changes (Windows)

**Tool call:** `word_live_replace_text`
```json
{
  "filename": "report.docx",
  "find_text": "ABC Corporation",
  "replace_text": "XYZ Ltd",
  "match_case": true,
  "replace_all": true,
  "track_changes": true
}
```
**Expected output:**
```json
{
  "status": "success",
  "replacements": 4,
  "message": "Replaced 4 occurrences (tracked changes enabled)"
}
```
The replacements appear as tracked changes in Word with strikethrough on "ABC Corporation" and underline on "XYZ Ltd".

### Example 3: Add a comment anchored to text (cross-platform)

**Tool call:** `add_comment`
```json
{
  "filename": "C:/Documents/contract.docx",
  "target_text": "payment within 30 days",
  "comment_text": "Should we extend this to 45 days?",
  "author": "Jane Smith"
}
```
**Expected output:**
```json
{
  "status": "success",
  "message": "Comment added by Jane Smith on 'payment within 30 days'"
}
```
The comment appears in Word's Review panel, anchored to the specified text.

## Tool Reference

**125 tools** across two modes — see the [complete tool reference](TOOLS.md) for details.

| Category | Count |
|----------|-------|
| Cross-platform (python-docx) | 80 |
| Windows Live (COM automation) | 45 |
| macOS Live (JXA automation) | 40 (of the 45 live tools) |

## Requirements

- **Python 3.11+**
- `python-docx`, `fastmcp`, `msoffcrypto-tool` (installed automatically)
- **Windows Live tools:** Windows 10/11 + Microsoft Word + `pywin32` (installed automatically)
- **macOS Live tools:** macOS + Microsoft Word for Mac (uses built-in JXA — no extra dependencies)

> The cross-platform tools work without Word installed — only python-docx is needed.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup, code style, and how to add new tools.

Found a bug? [Open an issue](https://github.com/GavinTomlins/word-mcp-live/issues/new?template=bug_report.md).
Have an idea? [Request a feature](https://github.com/GavinTomlins/word-mcp-live/issues/new?template=feature_request.md).

## Acknowledgments

This fork stands on two generations of prior work, gratefully acknowledged:

- **[GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server)** by GongRzhe — the original cross-platform Word MCP server this lineage descends from (MIT License).
- **[cheemscheems/word-mcp-live](https://github.com/cheemscheems/word-mcp-live)** by Yüce Karapazar — added the live editing engine (Windows COM and macOS JXA), native tracked changes, per-action undo, and the security hardening that this fork builds on.

This fork (GavinTomlins) contributes the FastMCP 3.x architecture, typed configuration, document validation, Markdown authoring/verification tools, and the bundled doc-oracle agent described in [What's New — GT Improvements](#whats-new--gt-improvements).

Additional libraries: [python-docx](https://python-docx.readthedocs.io/) &middot; [FastMCP](https://gofastmcp.com/) &middot; [pywin32](https://github.com/mhammond/pywin32)

## Privacy

This server runs entirely on your local machine. No data is collected, transmitted, or stored. See the full [Privacy Policy](PRIVACY.md).

## Support

- **Bug reports:** [Open an issue](https://github.com/GavinTomlins/word-mcp-live/issues/new?template=bug_report.md)
- **Feature requests:** [Request a feature](https://github.com/GavinTomlins/word-mcp-live/issues/new?template=feature_request.md)
- **Discussions:** [GitHub Discussions](https://github.com/GavinTomlins/word-mcp-live/discussions)

## License

MIT License — see [LICENSE](LICENSE) for details.

## Disclaimer

> **⚠️ This project is provided as-is, without any warranty or guarantee of fitness for a particular purpose.**
>
> While security issues identified in the initial audit have been addressed, this project has **not undergone comprehensive testing** in production environments. Use of this software to modify Word documents carries inherent risks, including but not limited to:
>
> - **Data loss or corruption** due to unexpected errors during file operations
> - **Document formatting issues** from automated editing
> - **Compatibility problems** with specific Word versions or document features
>
> **Always maintain backups** of important documents before using this tool. The automatic backup feature (stored in `_backup/` folders) provides a safety net, but should not be relied upon as your sole backup strategy.
>
> By using this software, you acknowledge that you understand and accept these risks.

<!-- mcp-name: io.github.gavintomlins/word-mcp-live -->
