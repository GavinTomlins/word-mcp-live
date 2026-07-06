"""Footnote and endnote tools.

Registrations moved verbatim from ``main.py``'s former ``register_tools()``.
"""

from mcp.types import ToolAnnotations

from word_mcp_live_gavintomlins.tools import footnote_tools


def register(mcp):
    """Register footnotes tools on ``mcp``."""
    # Footnote tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Footnote",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_footnote_to_document(filename: str, paragraph_index: int, footnote_text: str):
        """Add a footnote to a specific paragraph in a Word document."""
        return footnote_tools.add_footnote_to_document(filename, paragraph_index, footnote_text)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Footnote After Text",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_footnote_after_text(filename: str, search_text: str, footnote_text: str,
                               output_filename: str = None):
        """Add a footnote after specific text with proper superscript formatting.
        This enhanced function ensures footnotes display correctly as superscript."""
        return footnote_tools.add_footnote_after_text(filename, search_text, footnote_text, output_filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Footnote Before Text",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_footnote_before_text(filename: str, search_text: str, footnote_text: str,
                                output_filename: str = None):
        """Add a footnote before specific text with proper superscript formatting.
        This enhanced function ensures footnotes display correctly as superscript."""
        return footnote_tools.add_footnote_before_text(filename, search_text, footnote_text, output_filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Footnote Enhanced",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_footnote_enhanced(filename: str, paragraph_index: int, footnote_text: str,
                             output_filename: str = None):
        """Enhanced footnote addition with guaranteed superscript formatting.
        Adds footnote at the end of a specific paragraph with proper style handling."""
        return footnote_tools.add_footnote_enhanced(filename, paragraph_index, footnote_text, output_filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Endnote",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_endnote_to_document(filename: str, paragraph_index: int, endnote_text: str):
        """Add an endnote to a specific paragraph in a Word document."""
        return footnote_tools.add_endnote_to_document(filename, paragraph_index, endnote_text)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Customize Footnote Style",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def customize_footnote_style(filename: str, numbering_format: str = "1, 2, 3",
                                start_number: int = 1, font_name: str = None,
                                font_size: int = None):
        """Customize footnote numbering and formatting in a Word document."""
        return footnote_tools.customize_footnote_style(
            filename, numbering_format, start_number, font_name, font_size
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Delete Footnote",
            destructiveHint=True,
        ),
    )
    def delete_footnote_from_document(filename: str, footnote_id: int = None,
                                     search_text: str = None, output_filename: str = None):
        """Delete a footnote from a Word document.
        Identify the footnote either by ID (1, 2, 3, etc.) or by searching for text near it."""
        return footnote_tools.delete_footnote_from_document(
            filename, footnote_id, search_text, output_filename
        )

    # Robust footnote tools - Production-ready with comprehensive validation
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Footnote Robust",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_footnote_robust(filename: str, search_text: str = None,
                           paragraph_index: int = None, footnote_text: str = "",
                           validate_location: bool = True, auto_repair: bool = False):
        """Add footnote with robust validation and Word compliance.
        This is the production-ready version with comprehensive error handling."""
        return footnote_tools.add_footnote_robust_tool(
            filename, search_text, paragraph_index, footnote_text,
            validate_location, auto_repair
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Validate Footnotes",
            readOnlyHint=True,
        ),
    )
    def validate_document_footnotes(filename: str):
        """Validate all footnotes in document for coherence and compliance.
        Returns detailed report on ID conflicts, orphaned content, missing styles, etc."""
        return footnote_tools.validate_footnotes_tool(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Delete Footnote Robust",
            destructiveHint=True,
        ),
    )
    def delete_footnote_robust(filename: str, footnote_id: int = None,
                              search_text: str = None, clean_orphans: bool = True):
        """Delete footnote with comprehensive cleanup and orphan removal.
        Ensures complete removal from document.xml, footnotes.xml, and relationships."""
        return footnote_tools.delete_footnote_robust_tool(
            filename, footnote_id, search_text, clean_orphans
        )
