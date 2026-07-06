"""Extended document tools: paragraph text, find, highlights, PDF conversion.

Registrations moved verbatim from ``main.py``'s former ``register_tools()``.
"""

from mcp.types import ToolAnnotations

from word_mcp_live_gavintomlins.tools import extended_document_tools


def register(mcp):
    """Register extended tools on ``mcp``."""
    # Extended document tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get Paragraph Text",
            readOnlyHint=True,
        ),
    )
    def get_paragraph_text_from_document(filename: str, paragraph_index: int):
        """Get text from a specific paragraph in a Word document."""
        return extended_document_tools.get_paragraph_text_from_document(filename, paragraph_index)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Find Text",
            readOnlyHint=True,
        ),
    )
    def find_text_in_document(filename: str, text_to_find: str, match_case: bool = True,
                             whole_word: bool = False):
        """Find occurrences of specific text in a Word document."""
        return extended_document_tools.find_text_in_document(
            filename, text_to_find, match_case, whole_word
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get Highlighted Text",
            readOnlyHint=True,
        ),
    )
    def get_highlighted_text(filename: str, color: str = None):
        """Extract all highlighted/colored text from a Word document, including text inside tables."""
        return extended_document_tools.get_highlighted_text_from_document(filename, color)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Convert to PDF",
            destructiveHint=True,
        ),
    )
    def convert_to_pdf(filename: str, output_filename: str = None):
        """Convert a Word document to PDF format."""
        return extended_document_tools.convert_to_pdf(filename, output_filename)
