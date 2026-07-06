"""Modular MCP tool registration.

Each module exposes ``register(mcp)``; ``register_all`` wires the full set.
Live tools carry the ``live`` tag so platforms without Word COM/JXA can
disable them wholesale via ``mcp.disable(tags={"live"})``.
"""

from word_mcp_live_cheemscheems.mcp_tools import (
    comments,
    content,
    documents,
    extended,
    footnotes,
    formatting,
    guidance,
    hyperlinks,
    layout,
    live_edit,
    live_layout,
    live_read,
    protection,
    quality,
    tracked_changes,
)

_ALL_MODULES = (
    documents,
    content,
    formatting,
    protection,
    footnotes,
    extended,
    comments,
    hyperlinks,
    tracked_changes,
    layout,
    quality,
    live_edit,
    live_read,
    live_layout,
    guidance,
)


def register_all(mcp):
    """Register every tool module on the given FastMCP server."""
    for module in _ALL_MODULES:
        module.register(mcp)
