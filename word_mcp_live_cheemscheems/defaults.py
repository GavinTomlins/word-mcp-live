"""Default author identity for comments and tracked changes.

Sourced from the typed settings (env: WORD_MCP_AUTHOR / MCP_AUTHOR and
WORD_MCP_AUTHOR_INITIALS / MCP_AUTHOR_INITIALS).
"""

from word_mcp_live_cheemscheems.config import get_settings

_settings = get_settings()

DEFAULT_AUTHOR = _settings.author
DEFAULT_INITIALS = _settings.author_initials
