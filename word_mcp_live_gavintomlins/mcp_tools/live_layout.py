"""Live layout tools (Word must be running with the document open).

Registrations moved verbatim from ``main.py``'s former ``register_tools()``.
"""

from mcp.types import ToolAnnotations

from word_mcp_live_gavintomlins.tools import live_layout_tools

LIVE_TAGS = {"live"}


def register(mcp):
    """Register live layout tools on ``mcp``."""
    # --- Live layout tools (Windows only, requires Word running) ---

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Set Page Layout",
            destructiveHint=True,
        ),
    )
    def word_live_set_page_layout(
        filename: str = None,
        section_index: int = 1,
        orientation: str = None,
        page_width_inches: float = None,
        page_height_inches: float = None,
        margin_top_inches: float = None,
        margin_bottom_inches: float = None,
        margin_left_inches: float = None,
        margin_right_inches: float = None,
    ):
        """[Windows only] Set page layout (orientation, size, margins) for a section in a Word document open in Word. Requires Word running."""
        return live_layout_tools.word_live_set_page_layout(
            filename, section_index, orientation,
            page_width_inches, page_height_inches,
            margin_top_inches, margin_bottom_inches,
            margin_left_inches, margin_right_inches,
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Add Header/Footer",
            destructiveHint=True,
        ),
    )
    def word_live_add_header_footer(
        filename: str = None,
        section_index: int = 1,
        header_text: str = None,
        footer_text: str = None,
        header_alignment: str = "center",
        footer_alignment: str = "center",
    ):
        """[Windows only] Add header and/or footer to a section in a Word document open in Word. Requires Word running."""
        return live_layout_tools.word_live_add_header_footer(
            filename, section_index, header_text, footer_text,
            header_alignment, footer_alignment,
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Add Page Numbers",
            destructiveHint=True,
        ),
    )
    def word_live_add_page_numbers(
        filename: str = None,
        section_index: int = 1,
        position: str = "footer",
        alignment: str = "center",
        prefix: str = "",
        suffix: str = "",
        include_total: bool = False,
    ):
        """[Windows only] Add page numbers to header or footer in a Word document open in Word. Requires Word running."""
        return live_layout_tools.word_live_add_page_numbers(
            filename, section_index, position, alignment,
            prefix, suffix, include_total,
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Add Section Break",
            destructiveHint=True,
        ),
    )
    def word_live_add_section_break(
        filename: str = None,
        break_type: str = "new_page",
    ):
        """[Windows only] Add a section break (new_page, continuous, even_page, odd_page) to a Word document open in Word. Requires Word running."""
        return live_layout_tools.word_live_add_section_break(
            filename, break_type,
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Set Paragraph Spacing",
            destructiveHint=True,
        ),
    )
    def word_live_set_paragraph_spacing(
        filename: str = None,
        paragraph_index: int = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        space_before_pt: float = None,
        space_after_pt: float = None,
        line_spacing: float = None,
        line_spacing_rule: str = None,
        keep_with_next: bool = None,
        keep_together: bool = None,
        alignment: str = None,
    ):
        """[Windows only] Set paragraph spacing and layout properties in a Word document open in Word. Paragraphs are 1-indexed. Requires Word running."""
        return live_layout_tools.word_live_set_paragraph_spacing(
            filename, paragraph_index, start_paragraph, end_paragraph,
            space_before_pt, space_after_pt, line_spacing, line_spacing_rule,
            keep_with_next, keep_together, alignment,
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Add Bookmark",
            destructiveHint=True,
        ),
    )
    def word_live_add_bookmark(
        filename: str = None,
        paragraph_index: int = 1,
        bookmark_name: str = "",
    ):
        """[Windows only] Add a named bookmark at a paragraph in a Word document open in Word.
        Paragraph is 1-indexed. Requires Word running."""
        return live_layout_tools.word_live_add_bookmark(
            filename, paragraph_index, bookmark_name,
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Add Watermark",
            destructiveHint=True,
        ),
    )
    def word_live_add_watermark(
        filename: str = None,
        text: str = "TASLAK",
        font_size: int = 72,
        font_color: str = "C0C0C0",
        rotation: int = -45,
        section_index: int = 1,
    ):
        """[Windows only] Add a diagonal text watermark to a Word document open in Word. Requires Word running."""
        return live_layout_tools.word_live_add_watermark(
            filename, text, font_size, font_color, rotation, section_index,
        )
