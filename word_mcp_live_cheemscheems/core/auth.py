"""Bearer Token authentication for HTTP/SSE transport modes.

Usage::

    WORD_MCP_LIVE_API_KEY=sk-your-secret-key python -m word_mcp_live_cheemscheems

The key can also be placed in a ``.env`` file in the project root
(``load_dotenv()`` is called at startup in *main.py*).
"""

import os

from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request
from starlette.responses import JSONResponse

# Read from environment (or .env via load_dotenv in main.py).
# None = auth disabled (backwards-compatible with stdio / no-env setups).
WORD_MCP_LIVE_API_KEY: str | None = os.environ.get("WORD_MCP_LIVE_API_KEY")
_API_KEY_SET = bool(WORD_MCP_LIVE_API_KEY)


def is_auth_enabled() -> bool:
    """Return ``True`` if an API key is configured."""
    return _API_KEY_SET


class BearerTokenMiddleware(BaseHTTPMiddleware):
    """Starlette ASGI middleware validating ``Authorization: Bearer <key>``.

    Only applies when ``WORD_MCP_LIVE_API_KEY`` is set.
    """

    async def dispatch(self, request: Request, call_next):
        if not _API_KEY_SET:
            return await call_next(request)

        auth_header = request.headers.get("Authorization", "")
        expected = f"Bearer {WORD_MCP_LIVE_API_KEY}"

        if auth_header != expected:
            return JSONResponse(
                {"error": "未授权。请提供有效的 WORD_MCP_LIVE_API_KEY。"},
                status_code=401,
            )

        return await call_next(request)
