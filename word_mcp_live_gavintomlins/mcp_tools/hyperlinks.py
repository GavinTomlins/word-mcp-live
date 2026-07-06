"""Hyperlink management tools.

Registrations moved verbatim from ``main.py``'s former ``register_tools()``.
"""

from mcp.types import ToolAnnotations

from word_mcp_live_gavintomlins.tools import hyperlink_tools


def register(mcp):
    """Register hyperlinks tools on ``mcp``."""
    # Hyperlink tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Manage Hyperlinks",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def manage_hyperlinks(filename: str, action: str = "add", text: str = "",
                          url: str = "", paragraph_index: int = None):
        """Add or manage hyperlinks in a Word document.
        Finds the specified text and converts it to a clickable hyperlink with blue underline.

        Args:
            filename: Path to Word document
            action: Action to perform ("add" to add a hyperlink)
            text: Text to convert to a hyperlink
            url: URL the hyperlink should point to
            paragraph_index: If specified, only search in this paragraph (0-based)
        """
        return hyperlink_tools.manage_hyperlinks(filename, action, text, url, paragraph_index)
