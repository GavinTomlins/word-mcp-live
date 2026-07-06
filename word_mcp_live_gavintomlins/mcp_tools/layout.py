"""Cross-platform layout tools: page setup, headers/footers, watermarks.

Registrations moved verbatim from ``main.py``'s former ``register_tools()``.
"""

from mcp.types import ToolAnnotations

from word_mcp_live_gavintomlins.tools import layout_tools


def register(mcp):
    """Register layout tools on ``mcp``."""
    # --- Layout, header/footer, spacing, bookmark, watermark tools ---

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Set Page Layout",
            destructiveHint=True,
        ),
    )
    def set_page_layout(
        filename: str,
        section_index: int = 0,
        orientation: str = None,
        page_width_inches: float = None,
        page_height_inches: float = None,
        margin_top_inches: float = None,
        margin_bottom_inches: float = None,
        margin_left_inches: float = None,
        margin_right_inches: float = None,
    ):
        """Set page layout (orientation, size, margins) for a document section."""
        return layout_tools.set_page_layout(
            filename, section_index, orientation,
            page_width_inches, page_height_inches,
            margin_top_inches, margin_bottom_inches,
            margin_left_inches, margin_right_inches,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Header/Footer",
            destructiveHint=True,
        ),
    )
    def add_header_footer(
        filename: str,
        section_index: int = 0,
        header_text: str = None,
        footer_text: str = None,
        header_alignment: str = "center",
        footer_alignment: str = "center",
    ):
        """Add header and/or footer text to a document section."""
        return layout_tools.add_header_footer(
            filename, section_index, header_text, footer_text,
            header_alignment, footer_alignment,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Page Numbers",
            destructiveHint=True,
        ),
    )
    def add_page_numbers(
        filename: str,
        section_index: int = 0,
        position: str = "footer",
        alignment: str = "center",
        prefix: str = "",
        suffix: str = "",
        include_total: bool = False,
    ):
        """Add page numbers to header or footer using PAGE/NUMPAGES fields."""
        return layout_tools.add_page_numbers(
            filename, section_index, position, alignment,
            prefix, suffix, include_total,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Section Break",
            destructiveHint=True,
        ),
    )
    def add_section_break(filename: str, break_type: str = "new_page"):
        """Add a section break (new_page, continuous, even_page, odd_page)."""
        return layout_tools.add_section_break(filename, break_type)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Set Paragraph Spacing",
            destructiveHint=True,
        ),
    )
    def set_paragraph_spacing(
        filename: str,
        paragraph_index: int = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        space_before_pt: float = None,
        space_after_pt: float = None,
        line_spacing: float = None,
        line_spacing_rule: str = None,
    ):
        """Set paragraph spacing (before/after/line) for one or a range of paragraphs.
        line_spacing_rule: single, 1.5_lines, double, exactly, at_least, multiple."""
        return layout_tools.set_paragraph_spacing(
            filename, paragraph_index, start_paragraph, end_paragraph,
            space_before_pt, space_after_pt, line_spacing, line_spacing_rule,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Bookmark",
            destructiveHint=True,
        ),
    )
    def add_bookmark(filename: str, paragraph_index: int, bookmark_name: str):
        """Add a named bookmark at a paragraph for cross-referencing."""
        return layout_tools.add_bookmark(filename, paragraph_index, bookmark_name)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Watermark",
            destructiveHint=True,
        ),
    )
    def add_watermark(
        filename: str,
        text: str = "TASLAK",
        font_size: int = 72,
        font_color: str = "C0C0C0",
        rotation: int = -45,
        section_index: int = 0,
    ):
        """Add a diagonal text watermark (e.g. TASLAK, GİZLİ, DRAFT) to a document."""
        return layout_tools.add_watermark(
            filename, text, font_size, font_color, rotation, section_index,
        )
