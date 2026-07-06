"""FastMCP server factory for the Word Document MCP Server.

``build_server`` constructs a fully configured server: instructions,
authentication, lifespan (save/path hooks + backup loop), per-tool logging
middleware, health route, and platform-based filtering of live tools.
"""

import logging
import sys
import time
from contextlib import asynccontextmanager

from fastmcp import FastMCP
from fastmcp.server.middleware import Middleware, MiddlewareContext
from starlette.requests import Request
from starlette.responses import JSONResponse

from word_mcp_live_cheemscheems.config import Settings, get_settings
from word_mcp_live_cheemscheems.core.auth import build_token_verifier
from word_mcp_live_cheemscheems.mcp_tools import register_all

log = logging.getLogger(__name__)

INSTRUCTIONS = (
    "Microsoft Word document editing via MCP. Two modes:\n"
    "1. Cross-platform tools (no prefix): require the .docx file to be CLOSED in Word.\n"
    '2. Live editing tools (prefix "word_live_"): require the file to be OPEN in Word.\n'
    "If a cross-platform tool fails with a file lock error, switch to the "
    "corresponding word_live_* tool instead."
)

# Platforms with a Word automation bridge (COM on Windows, JXA on macOS).
LIVE_PLATFORMS = ("win32", "darwin")


class ToolLoggingMiddleware(Middleware):
    """Log every tool call with name, latency and outcome."""

    async def on_call_tool(self, context: MiddlewareContext, call_next):
        tool = getattr(context.message, "name", "<unknown>")
        start = time.perf_counter()
        try:
            result = await call_next(context)
        except Exception as exc:
            log.warning(
                "tool %s failed after %.1f ms: %s",
                tool,
                (time.perf_counter() - start) * 1000,
                exc,
            )
            raise
        log.info(
            "tool %s completed in %.1f ms",
            tool,
            (time.perf_counter() - start) * 1000,
        )
        return result


@asynccontextmanager
async def _lifespan(server):
    """Install document hooks and run the backup loop for the server's lifetime."""
    from word_mcp_live_cheemscheems.utils.backup_manager import backup_manager
    from word_mcp_live_cheemscheems.utils.path_utils import install_path_hook
    from word_mcp_live_cheemscheems.utils.save_utils import install_save_hook

    # Monkey-patch Document.save() to preserve comments.xml and other
    # custom parts, and PhysPkgReader to detect Word-locked files.
    install_save_hook()
    install_path_hook()
    backup_manager.start()
    try:
        yield
    finally:
        backup_manager.stop()


def build_server(
    settings: Settings | None = None,
    platform: str = sys.platform,
) -> FastMCP:
    """Construct the FastMCP server with all tools registered."""
    settings = settings or get_settings()

    mcp = FastMCP(
        name="Word Document Server",
        instructions=INSTRUCTIONS,
        lifespan=_lifespan,
        auth=build_token_verifier(settings) if settings.transport != "stdio" else None,
        middleware=[ToolLoggingMiddleware()],
        on_duplicate="error",
        mask_error_details=settings.mask_error_details,
    )

    register_all(mcp)

    if platform not in LIVE_PLATFORMS:
        mcp.disable(tags={"live"})
        log.info(
            "live editing tools disabled: platform %r has no Word COM/JXA bridge",
            platform,
        )

    @mcp.custom_route("/health", methods=["GET"])
    async def health(request: Request) -> JSONResponse:
        return JSONResponse({"status": "ok", "server": "word-mcp-live"})

    return mcp
