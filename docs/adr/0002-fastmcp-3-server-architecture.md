# ADR 0002: FastMCP 3.x server architecture

- **Status:** Accepted (retrospective — records the July 2026 restructure)
- **Date:** 2026-07-06

## Context

The server was written against FastMCP 2.x idioms while its lockfile resolved
FastMCP 3.4.3, which rejected the `description=` constructor kwarg at import
time — the server could not start. Registration of ~120 tools lived in one
1,900-line function; configuration was scattered `os.getenv` calls; HTTP auth
was a hand-rolled ASGI middleware; the deprecated SSE transport was still
offered; startup side effects (monkey-patches, backup loop) ran imperatively
with no teardown.

## Decision

Align with the architecture recommended by the FastMCP server documentation
(gofastmcp.com/servers/server), using a sibling project
(openproject-gt-mcp-server) as the reference implementation:

1. **Factory construction** — `build_server(settings)` in `server.py` builds
   the `FastMCP` instance with `instructions`, `on_duplicate="error"`, and
   `mask_error_details` (masked by default on HTTP, full detail on stdio).
2. **Typed settings** — `config.py` (pydantic-settings) with `WORD_MCP_*`
   names and aliases for every legacy variable, so existing deployments keep
   working.
3. **Native auth** — `StaticTokenVerifier` via the `auth=` constructor
   parameter replaces the custom bearer middleware.
4. **Lifespan** — save/path monkey-patches and the backup loop start inside
   the FastMCP lifespan and stop on shutdown.
5. **Modular registration** — `mcp_tools/` package, one module per tool
   category, each exposing `register(mcp)`.
6. **Tag-based platform gating** — live tools carry the `live` tag;
   `mcp.disable(tags={"live"})` hides them on platforms without a Word
   COM/JXA bridge instead of letting them fail at call time.
7. **Transports** — stdio and streamable HTTP only; SSE removed.

## Process

The 2.x/3.x mismatch was found by inspecting the constructor signature of the
locked FastMCP version and reproducing the `TypeError` in the project venv.
The target architecture was derived by diffing this server against
openproject-gt-mcp-server and the FastMCP docs, then verified with an
in-memory FastMCP client (tool counts per platform, auth behaviour over ASGI,
lifespan start/stop) before merge.

## Consequences

- The server starts on FastMCP 3.x on all three platforms.
- Tool modules are independently testable; duplicate registrations fail fast.
- Clients on Linux never see tools that cannot work there.
- SSE users must move to streamable HTTP (`WORD_MCP_TRANSPORT=http`).
