"""Formatting tools: styles, text/table formatting, cell shading and widths.

Registrations moved verbatim from ``main.py``'s former ``register_tools()``.
"""

from mcp.types import ToolAnnotations

from word_mcp_live_gavintomlins.tools import format_tools


def register(mcp):
    """Register formatting tools on ``mcp``."""
    # Format tools (styling, text formatting, etc.)
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Create Custom Style",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def create_custom_style(filename: str, style_name: str, bold: bool = None,
                          italic: bool = None, font_size: int = None,
                          font_name: str = None, color: str = None,
                          base_style: str = None):
        """Create a custom style in the document."""
        return format_tools.create_custom_style(
            filename, style_name, bold, italic, font_size, font_name, color, base_style
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Format Text",
            readOnlyHint=False,
            destructiveHint=False,
        ),
        description=format_tools.format_text.__doc__,
    )
    def format_text(filename: str, paragraph_index: int, start_pos: int, end_pos: int,
                   bold: bool = None, italic: bool = None, underline: bool = None,
                   color: str = None, font_size: int = None, font_name: str = None):
        return format_tools.format_text(
            filename, paragraph_index, start_pos, end_pos, bold, italic,
            underline, color, font_size, font_name
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Format Table",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def format_table(filename: str, table_index: int, has_header_row: bool = None,
                    border_style: str = None, shading: list[str] = None):
        """Format a table with borders, shading, and structure."""
        return format_tools.format_table(filename, table_index, has_header_row, border_style, shading)

    # New table cell shading tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Set Table Cell Shading",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def set_table_cell_shading(filename: str, table_index: int, row_index: int,
                              col_index: int, fill_color: str, pattern: str = "clear"):
        """Apply shading/filling to a specific table cell."""
        return format_tools.set_table_cell_shading(filename, table_index, row_index, col_index, fill_color, pattern)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Apply Alternating Row Colors",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def apply_table_alternating_rows(filename: str, table_index: int,
                                   color1: str = "FFFFFF", color2: str = "F2F2F2"):
        """Apply alternating row colors to a table for better readability."""
        return format_tools.apply_table_alternating_rows(filename, table_index, color1, color2)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Highlight Table Header",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def highlight_table_header(filename: str, table_index: int,
                             header_color: str = "4472C4", text_color: str = "FFFFFF"):
        """Apply special highlighting to table header row."""
        return format_tools.highlight_table_header(filename, table_index, header_color, text_color)

    # Cell merging tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Merge Table Cells",
            readOnlyHint=False,
            destructiveHint=True,
        ),
    )
    def merge_table_cells(filename: str, table_index: int, start_row: int, start_col: int,
                        end_row: int, end_col: int):
        """Merge cells in a rectangular area of a table."""
        return format_tools.merge_table_cells(filename, table_index, start_row, start_col, end_row, end_col)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Merge Cells Horizontally",
            readOnlyHint=False,
            destructiveHint=True,
        ),
    )
    def merge_table_cells_horizontal(filename: str, table_index: int, row_index: int,
                                   start_col: int, end_col: int):
        """Merge cells horizontally in a single row."""
        return format_tools.merge_table_cells_horizontal(filename, table_index, row_index, start_col, end_col)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Merge Cells Vertically",
            readOnlyHint=False,
            destructiveHint=True,
        ),
    )
    def merge_table_cells_vertical(filename: str, table_index: int, col_index: int,
                                 start_row: int, end_row: int):
        """Merge cells vertically in a single column."""
        return format_tools.merge_table_cells_vertical(filename, table_index, col_index, start_row, end_row)

    # Cell alignment tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Set Cell Alignment",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def set_table_cell_alignment(filename: str, table_index: int, row_index: int, col_index: int,
                               horizontal: str = "left", vertical: str = "top"):
        """Set text alignment for a specific table cell."""
        return format_tools.set_table_cell_alignment(filename, table_index, row_index, col_index, horizontal, vertical)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Set Table Alignment",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def set_table_alignment_all(filename: str, table_index: int,
                              horizontal: str = "left", vertical: str = "top"):
        """Set text alignment for all cells in a table."""
        return format_tools.set_table_alignment_all(filename, table_index, horizontal, vertical)


    # New table column width tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Set Column Width",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def set_table_column_width(filename: str, table_index: int, col_index: int,
                              width: float, width_type: str = "points"):
        """Set the width of a specific table column."""
        return format_tools.set_table_column_width(filename, table_index, col_index, width, width_type)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Set Column Widths",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def set_table_column_widths(filename: str, table_index: int, widths: list[float],
                               width_type: str = "points"):
        """Set the widths of multiple table columns."""
        return format_tools.set_table_column_widths(filename, table_index, widths, width_type)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Set Table Width",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def set_table_width(filename: str, table_index: int, width: float,
                       width_type: str = "points"):
        """Set the overall width of a table."""
        return format_tools.set_table_width(filename, table_index, width, width_type)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Auto-Fit Table Columns",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def auto_fit_table_columns(filename: str, table_index: int):
        """Set table columns to auto-fit based on content."""
        return format_tools.auto_fit_table_columns(filename, table_index)

    # New table cell text formatting and padding tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Format Cell Text",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def format_table_cell_text(filename: str, table_index: int, row_index: int, col_index: int,
                               text_content: str = None, bold: bool = None, italic: bool = None,
                               underline: bool = None, color: str = None, font_size: int = None,
                               font_name: str = None):
        """Format text within a specific table cell."""
        return format_tools.format_table_cell_text(filename, table_index, row_index, col_index,
                                                   text_content, bold, italic, underline, color, font_size, font_name)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Set Cell Padding",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def set_table_cell_padding(filename: str, table_index: int, row_index: int, col_index: int,
                               top: float = None, bottom: float = None, left: float = None,
                               right: float = None, unit: str = "points"):
        """Set padding/margins for a specific table cell."""
        return format_tools.set_table_cell_padding(filename, table_index, row_index, col_index,
                                                   top, bottom, left, right, unit)
