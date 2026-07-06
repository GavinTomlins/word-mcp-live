"""Comment tools: read and write document comments.

Registrations moved verbatim from ``main.py``'s former ``register_tools()``.
"""

from mcp.types import ToolAnnotations

from word_mcp_live_cheemscheems.defaults import DEFAULT_AUTHOR, DEFAULT_INITIALS
from word_mcp_live_cheemscheems.tools import comment_tools, comment_write_tools


def register(mcp):
    """Register comments tools on ``mcp``."""
    # Comment tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get All Comments",
            readOnlyHint=True,
        ),
    )
    def get_all_comments(filename: str):
        """Extract all comments from a Word document."""
        return comment_tools.get_all_comments(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get Comments by Author",
            readOnlyHint=True,
        ),
    )
    def get_comments_by_author(filename: str, author: str):
        """Extract comments from a specific author in a Word document."""
        return comment_tools.get_comments_by_author(filename, author)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get Comments for Paragraph",
            readOnlyHint=True,
        ),
    )
    def get_comments_for_paragraph(filename: str, paragraph_index: int):
        """Extract comments for a specific paragraph in a Word document."""
        return comment_tools.get_comments_for_paragraph(filename, paragraph_index)
    # Comment write tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Comment",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_comment(filename: str, target_text: str, comment_text: str,
                    author: str = DEFAULT_AUTHOR, initials: str = DEFAULT_INITIALS):
        """Add a comment to a Word document anchored to specific text.

        [cross-platform mode] Requires the .docx file to be CLOSED in Word.
        If the file is open in Word, use word_live_add_comment instead.

        The comment will appear in Word's Review panel attached to the target text.

        Args:
            filename: Path to Word document
            target_text: Text in the document to attach the comment to
            comment_text: The comment content
            author: Comment author name (env: MCP_AUTHOR)
            initials: Author initials (env: MCP_AUTHOR_INITIALS)
        """
        return comment_write_tools.add_comment(filename, target_text, comment_text, author, initials)
