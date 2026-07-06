"""Tracked-changes tools: track edits, list/accept/reject revisions.

Registrations moved verbatim from ``main.py``'s former ``register_tools()``.
"""

from mcp.types import ToolAnnotations

from word_mcp_live_cheemscheems.defaults import DEFAULT_AUTHOR
from word_mcp_live_cheemscheems.tools import tracked_changes_tools


def register(mcp):
    """Register tracked changes tools on ``mcp``."""
    # Tracked changes tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Track Replace",
            destructiveHint=True,
        ),
        description=tracked_changes_tools.track_replace.__doc__,
    )
    def track_replace(filename: str, old_text: str, new_text: str, author: str = DEFAULT_AUTHOR):
        return tracked_changes_tools.track_replace(filename, old_text, new_text, author)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Track Insert",
            destructiveHint=True,
        ),
        description=tracked_changes_tools.track_insert.__doc__,
    )
    def track_insert(filename: str, after_text: str, insert_text: str, author: str = DEFAULT_AUTHOR):
        return tracked_changes_tools.track_insert(filename, after_text, insert_text, author)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Track Delete",
            destructiveHint=True,
        ),
        description=tracked_changes_tools.track_delete.__doc__,
    )
    def track_delete(filename: str, text: str, author: str = DEFAULT_AUTHOR):
        return tracked_changes_tools.track_delete(filename, text, author)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="List Tracked Changes",
            readOnlyHint=True,
        ),
    )
    def list_tracked_changes(filename: str):
        """List all tracked changes (insertions and deletions) in a Word document.
        Returns author, date, text, and paragraph context for each change."""
        return tracked_changes_tools.list_tracked_changes(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Accept Tracked Changes",
            destructiveHint=True,
        ),
    )
    def accept_tracked_changes(filename: str, author: str = None, change_ids: list[int] = None):
        """Accept tracked changes: apply insertions (keep text) and remove deletions.
        Optionally filter by author or specific change IDs."""
        return tracked_changes_tools.accept_tracked_changes(filename, author, change_ids)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Reject Tracked Changes",
            destructiveHint=True,
        ),
    )
    def reject_tracked_changes(filename: str, author: str = None, change_ids: list[int] = None):
        """Reject tracked changes: remove insertions and restore deleted text.
        Optionally filter by author or specific change IDs."""
        return tracked_changes_tools.reject_tracked_changes(filename, author, change_ids)
