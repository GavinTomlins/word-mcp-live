"""Live editing tools (Word must be running with the document open).

Registrations moved verbatim from ``main.py``'s former ``register_tools()``.
"""

from mcp.types import ToolAnnotations

from word_mcp_live_cheemscheems.tools import live_tools, screen_capture_tools

LIVE_TAGS = {"live"}


def register(mcp):
    """Register live edit tools on ``mcp``."""
    # --- Live editing tools (Windows only, requires Word running) ---

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Screen Capture",
            readOnlyHint=True,
        ),
    )
    def word_screen_capture(filename: str = None, output_path: str = None):
        """[Windows only] Capture a screenshot of a Word document window.
        Returns the path to the saved PNG image. Requires Word to be running."""
        return screen_capture_tools.word_screen_capture(filename, output_path)

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Insert Text",
            destructiveHint=True,
        ),
    )
    def word_live_insert_text(
        filename: str = None,
        text: str = "",
        position: str = "end",
        bookmark: str = None,
        track_changes: bool = False,
    ):
        """[Windows only] Insert text into a Word document that is open in Word.
        Position: 'start', 'end', 'cursor', or character offset. Requires Word running."""
        return live_tools.word_live_insert_text(
            filename, text, position, bookmark, track_changes
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Format Text",
            destructiveHint=True,
        ),
        description=live_tools.word_live_format_text.__doc__,
    )
    def word_live_format_text(
        filename: str = None,
        start: int = None,
        end: int = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        bold: bool = None,
        italic: bool = None,
        underline: bool = None,
        strikethrough: bool = None,
        font_name: str = None,
        font_size: float = None,
        font_color: str = None,
        highlight_color: int = None,
        style_name: str = None,
        paragraph_alignment: str = None,
        page_break_before: bool = None,
        preserve_direct_formatting: bool = False,
        track_changes: bool = False,
    ):
        return live_tools.word_live_format_text(
            filename, start, end, start_paragraph, end_paragraph,
            bold, italic, underline, strikethrough,
            font_name, font_size, font_color, highlight_color,
            style_name, paragraph_alignment, page_break_before,
            preserve_direct_formatting, track_changes,
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Replace Text",
            destructiveHint=True,
        ),
        description=live_tools.word_live_replace_text.__doc__,
    )
    def word_live_replace_text(
        filename: str = None,
        find_text: str = "",
        replace_text: str = "",
        match_case: bool = False,
        match_whole_word: bool = False,
        use_wildcards: bool = False,
        replace_all: bool = True,
        track_changes: bool = False,
    ):
        return live_tools.word_live_replace_text(
            filename, find_text, replace_text, match_case,
            match_whole_word, use_wildcards, replace_all, track_changes,
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Insert Paragraphs",
            destructiveHint=True,
        ),
        description=live_tools.word_live_insert_paragraphs.__doc__,
    )
    def word_live_insert_paragraphs(
        filename: str = None,
        paragraphs: list = None,
        target_text: str = None,
        target_paragraph_index: int = None,
        position: str = "after",
        style: str = None,
        track_changes: bool = False,
    ):
        return live_tools.word_live_insert_paragraphs(
            filename, paragraphs, target_text, target_paragraph_index,
            position, style, track_changes,
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Add Table",
            destructiveHint=True,
        ),
    )
    def word_live_add_table(
        filename: str = None,
        rows: int = 2,
        cols: int = 2,
        position: str = "end",
        data: list = None,
        style: str = "Table Grid",
        autofit: str = "window",
        track_changes: bool = False,
    ):
        """[live editing mode] Add a table to a document currently OPEN in Word.
        Use this when add_table fails due to file lock.
        Optionally provide data as 2D list. Default style is 'Table Grid' with
        autofit to window width. Requires Word running."""
        return live_tools.word_live_add_table(
            filename, rows, cols, position, data, style, autofit, track_changes
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Format Table",
            destructiveHint=True,
        ),
        description=live_tools.word_live_format_table.__doc__,
    )
    def word_live_format_table(
        filename: str = None,
        table_index: int = -1,
        border_style: str = None,
        cell_bold: list = None,
        cell_alignment: list = None,
        column_widths: list = None,
        table_alignment: str = None,
        cell_shading: list = None,
        autofit: str = None,
    ):
        return live_tools.word_live_format_table(
            filename, table_index, border_style, cell_bold, cell_alignment,
            column_widths, table_alignment, cell_shading, autofit
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Modify Table",
            destructiveHint=True,
        ),
        description=live_tools.word_live_modify_table.__doc__,
    )
    def word_live_modify_table(
        filename: str = None,
        table_index: int = 1,
        operation: str = "get_info",
        row: int = None,
        col: int = None,
        text: str = None,
        before_row: int = None,
        before_col: int = None,
        header: str = None,
        cells: list = None,
        start_row: int = None,
        start_col: int = None,
        end_row: int = None,
        end_col: int = None,
        autofit_mode: str = "content",
        accept_revisions: bool = False,
        track_changes: bool = False,
    ):
        return live_tools.word_live_modify_table(
            filename, table_index, operation, row, col, text,
            before_row, before_col, header, cells,
            start_row, start_col, end_row, end_col,
            autofit_mode, accept_revisions, track_changes,
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Delete Text",
            destructiveHint=True,
        ),
    )
    def word_live_delete_text(
        filename: str = None,
        start: int = None,
        end: int = None,
        track_changes: bool = False,
    ):
        """[Windows only] Delete text from a Word document open in Word.
        Specify start/end character positions. Requires Word running."""
        return live_tools.word_live_delete_text(
            filename, start, end, track_changes
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Apply List",
            destructiveHint=True,
        ),
        description=live_tools.word_live_apply_list.__doc__,
    )
    def word_live_apply_list(
        filename: str = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        list_type: str = "bullet",
        level: int = 0,
        remove: bool = False,
        continue_previous: bool = False,
        number_format: dict = None,
        number_style: dict = None,
        start_at: dict = None,
        level_map: dict = None,
        track_changes: bool = False,
        font_color: str = None,
        outline_numbered: bool = True,
    ):
        return live_tools.word_live_apply_list(
            filename, start_paragraph, end_paragraph, list_type,
            level, remove, continue_previous, number_format,
            number_style, start_at, level_map, track_changes,
            font_color, outline_numbered,
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Setup Heading Numbering",
            destructiveHint=True,
        ),
        description=live_tools.word_live_setup_heading_numbering.__doc__,
    )
    def word_live_setup_heading_numbering(
        filename: str = None,
        h1_paragraphs: list = None,
        h2_paragraphs: list = None,
        heading_map: dict = None,
        strip_manual_numbers: bool = True,
        h1_number_format: str = None,
        h2_number_format: str = None,
        font_name: str = None,
        h1_size: float = None,
        h2_size: float = None,
        bold: bool = None,
        alignment: str = None,
        font_color: str = None,
        h1_space_before: float = None,
        h1_space_after: float = None,
        h2_space_before: float = None,
        h2_space_after: float = None,
        line_spacing: float = None,
    ):
        return live_tools.word_live_setup_heading_numbering(
            filename, h1_paragraphs, h2_paragraphs, heading_map,
            strip_manual_numbers,
            h1_number_format, h2_number_format,
            font_name, h1_size, h2_size, bold, alignment, font_color,
            h1_space_before, h1_space_after, h2_space_before, h2_space_after,
            line_spacing,
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Undo",
            destructiveHint=True,
        ),
    )
    def word_live_undo(
        filename: str = None,
        times: int = 1,
    ):
        """[Windows only] Undo the last N operations in a Word document open in Word.
        Each MCP tool call is one undo entry. Requires Word running."""
        return live_tools.word_live_undo(filename, times)

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Save",
            destructiveHint=True,
        ),
    )
    def word_live_save(
        filename: str = None,
        save_as: str = None,
    ):
        """[Windows only] Save a Word document open in Word.
        Optionally save to a new path with save_as. Requires Word running."""
        return live_tools.word_live_save(filename, save_as)

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Toggle Track Changes",
            destructiveHint=True,
        ),
    )
    def word_live_toggle_track_changes(
        filename: str = None,
        enable: bool = None,
    ):
        """[Windows only] Toggle or set Track Changes mode on a Word document.
        If enable is omitted, toggles current state. Requires Word running."""
        return live_tools.word_live_toggle_track_changes(filename, enable)

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Insert Image",
            destructiveHint=True,
        ),
        description=live_tools.word_live_insert_image.__doc__,
    )
    def word_live_insert_image(
        filename: str = None,
        image_path: str = "",
        paragraph_index: int = None,
        position: str = "end",
        width_inches: float = None,
        height_inches: float = None,
        width_pt: float = None,
        height_pt: float = None,
        alignment: str = None,
        wrapping: str = None,
        border_style: str = None,
        border_width_pt: float = None,
        border_color: str = None,
        link_to_file: bool = False,
    ):
        return live_tools.word_live_insert_image(
            filename, image_path, paragraph_index, position,
            width_inches, height_inches, width_pt, height_pt,
            alignment, wrapping, border_style, border_width_pt,
            border_color, link_to_file
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Insert Cross Reference",
            destructiveHint=True,
        ),
    )
    def word_live_insert_cross_reference(
        filename: str = None,
        ref_type: str = "heading",
        ref_item: int = 1,
        ref_kind: str = "text",
        insert_position: str = "end",
        paragraph_index: int = None,
        insert_as_hyperlink: bool = True,
    ):
        """[Windows only] Insert a cross-reference to a heading, bookmark, figure, table, etc.
        First use word_live_list_cross_reference_items to discover available targets.
        ref_type: heading, bookmark, figure, table, equation, footnote, endnote.
        ref_kind: text, number, number_no_context, page, above_below.
        Requires Word running."""
        return live_tools.word_live_insert_cross_reference(
            filename, ref_type, ref_item, ref_kind,
            insert_position, paragraph_index, insert_as_hyperlink
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live List Cross Reference Items",
            readOnlyHint=True,
        ),
    )
    def word_live_list_cross_reference_items(
        filename: str = None,
        ref_type: str = "heading",
    ):
        """[Windows only] List available cross-reference targets in a Word document.
        Returns items that can be referenced with word_live_insert_cross_reference.
        ref_type: heading, bookmark, figure, table, equation, footnote, endnote.
        Requires Word running."""
        return live_tools.word_live_list_cross_reference_items(filename, ref_type)

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Insert Equation",
            destructiveHint=True,
        ),
    )
    def word_live_insert_equation(
        filename: str = None,
        equation: str = "",
        paragraph_index: int = None,
        position: str = "end",
        display_mode: bool = False,
    ):
        """[Windows only] Insert a mathematical equation into a Word document using UnicodeMath syntax.
        Examples: "x^2 + y^2 = z^2", "(a+b)/(c+d)" (fraction), "\\sqrt(x^2+y^2)" (root),
        "\\alpha + \\beta" (Greek), "\\int_0^\\infty e^(-x^2) dx" (integral),
        "\\sum_(i=1)^n i^2" (summation), "\\matrix(a&b@c&d)" (matrix).
        display_mode=True centers the equation on its own line.
        Requires Word running."""
        return live_tools.word_live_insert_equation(
            filename, equation, paragraph_index, position, display_mode
        )
