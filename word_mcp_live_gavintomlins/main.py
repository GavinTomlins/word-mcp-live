"""
Main entry point for the Word Document MCP Server.

Builds the FastMCP server (see ``server.py``) and runs it on the configured
transport: stdio (default) or streamable HTTP. Tool registrations live in
the ``mcp_tools`` package; configuration in ``config.py``.
"""

import logging
import sys

from dotenv import load_dotenv


def run_server():
    """Run the Word Document MCP Server with configurable transport."""
    load_dotenv()

    from pydantic import ValidationError

    from word_mcp_live_gavintomlins.config import get_settings

    try:
        settings = get_settings()
    except ValidationError as e:
        print(f"Invalid configuration: {e}", file=sys.stderr)
        sys.exit(1)

    logging.basicConfig(
        level=getattr(logging, settings.log_level.upper(), logging.INFO),
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        stream=sys.stderr,
    )
    log = logging.getLogger(__name__)

    from word_mcp_live_gavintomlins.server import build_server

    mcp = build_server(settings)

    try:
        if settings.transport == "stdio":
            log.info("starting on stdio transport")
            mcp.run(transport="stdio", show_banner=False)
            return mcp

        # HTTP transport: authentication is required unless explicitly waived.
        if settings.auth_required:
            print(
                "ERROR: an API key is required for the HTTP transport.\n"
                "  To enable authentication:\n"
                '    export WORD_MCP_API_KEY="your-secret-key"\n'
                "    (or add WORD_MCP_API_KEY=your-secret-key to .env)\n"
                "  To disable auth (local/dev only):\n"
                "    export WORD_MCP_INSECURE=true",
                file=sys.stderr,
            )
            sys.exit(1)

        if not settings.api_key:
            log.warning(
                "running in insecure mode (no authentication); "
                "set WORD_MCP_API_KEY for production use"
            )

        log.info(
            "starting on http transport at http://%s:%s%s",
            settings.host,
            settings.port,
            settings.path,
        )
        mcp.run(
            transport="http",
            host=settings.host,
            port=settings.port,
            path=settings.path,
            show_banner=False,
        )
    except KeyboardInterrupt:
        print("\nShutting down server...", file=sys.stderr)
    except Exception:
        logging.getLogger(__name__).exception("Error starting server")
        sys.exit(1)


if __name__ == "__main__":
    run_server()
