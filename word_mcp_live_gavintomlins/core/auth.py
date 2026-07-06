"""Bearer-token authentication for the HTTP transport.

Auth uses FastMCP's native ``auth=`` mechanism: a ``StaticTokenVerifier``
that accepts the single configured API key, giving standard 401 handling
(including ``WWW-Authenticate``) for free.

Usage::

    # Require authentication (default for HTTP):
    WORD_MCP_API_KEY=sk-your-secret-key python -m word_mcp_live_gavintomlins

    # Explicitly allow no-auth (NOT RECOMMENDED for remote access):
    WORD_MCP_INSECURE=true python -m word_mcp_live_gavintomlins

The legacy names ``WORD_MCP_LIVE_API_KEY`` / ``WORD_MCP_LIVE_INSECURE``
still work; the key can also live in a ``.env`` file in the project root.
"""

from fastmcp.server.auth.providers.jwt import StaticTokenVerifier

from word_mcp_live_gavintomlins.config import Settings


def build_token_verifier(settings: Settings) -> StaticTokenVerifier | None:
    """Return a verifier for the configured API key, or None when auth is off.

    The caller is responsible for refusing to serve HTTP when
    ``settings.auth_required`` is True (no key and no insecure opt-in).
    """
    if not settings.api_key:
        return None
    return StaticTokenVerifier(
        tokens={settings.api_key: {"client_id": "word-mcp-live"}},
    )
