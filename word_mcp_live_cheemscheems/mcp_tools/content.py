"""Content tools: paragraphs, headings, tables, lists, TOC, search/replace.

Registrations moved verbatim from ``main.py``'s former ``register_tools()``.
"""

from mcp.types import ToolAnnotations

from word_mcp_live_cheemscheems.tools import content_tools
from word_mcp_live_cheemscheems.tools.content_tools import (
    replace_block_between_manual_anchors_tool,
    replace_paragraph_block_below_header_tool,
)


def register(mcp):
    """Register content tools on ``mcp``."""
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Insert Header Near Text",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def insert_header_near_text(filename: str, target_text: str = None, header_title: str = None, position: str = 'after', header_style: str = 'Heading 1', target_paragraph_index: int = None):
        """Insert a header (with specified style) before or after the target paragraph. Specify by text or paragraph index. Args: filename (str), target_text (str, optional), header_title (str), position ('before' or 'after'), header_style (str, default 'Heading 1'), target_paragraph_index (int, optional)."""
        return content_tools.insert_header_near_text_tool(filename, target_text, header_title, position, header_style, target_paragraph_index)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Insert Line Near Text",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def insert_line_or_paragraph_near_text(filename: str, target_text: str = None, line_text: str = None, position: str = 'after', line_style: str = None, target_paragraph_index: int = None):
        """
        Insert a new line or paragraph (with specified or matched style) before or after the target paragraph. Specify by text or paragraph index. Args: filename (str), target_text (str, optional), line_text (str), position ('before' or 'after'), line_style (str, optional), target_paragraph_index (int, optional).
        """
        return content_tools.insert_line_or_paragraph_near_text_tool(filename, target_text, line_text, position, line_style, target_paragraph_index)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Insert List Near Text",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def insert_numbered_list_near_text(filename: str, target_text: str = None, list_items: list[str] = None, position: str = 'after', target_paragraph_index: int = None, bullet_type: str = 'bullet'):
        """Insert a bulleted or numbered list before or after the target paragraph. Specify by text or paragraph index. Args: filename (str), target_text (str, optional), list_items (list of str), position ('before' or 'after'), target_paragraph_index (int, optional), bullet_type ('bullet' for bullets or 'number' for numbered lists, default: 'bullet')."""
        return content_tools.insert_numbered_list_near_text_tool(filename, target_text, list_items, position, target_paragraph_index, bullet_type)
    # Content tools (paragraphs, headings, tables, etc.)
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Paragraph",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_paragraph(filename: str, text: str, style: str = None,
                      font_name: str = None, font_size: int = None,
                      bold: bool = None, italic: bool = None, color: str = None):
        """Add a paragraph to a Word document with optional formatting.

        [cross-platform mode] Requires the .docx file to be CLOSED in Word.
        If the file is open in Word, use word_live_insert_paragraphs instead.

        Args:
            filename: Path to Word document
            text: Paragraph text content
            style: Optional paragraph style name
            font_name: Font family (e.g., 'Helvetica', 'Times New Roman')
            font_size: Font size in points (e.g., 14, 36)
            bold: Make text bold
            italic: Make text italic
            color: Text color as hex RGB (e.g., '000000')
        """
        return content_tools.add_paragraph(filename, text, style, font_name, font_size, bold, italic, color)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Heading",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_heading(filename: str, text: str, level: int = 1,
                    font_name: str = None, font_size: int = None,
                    bold: bool = None, italic: bool = None, border_bottom: bool = False):
        """Add a heading to a Word document with optional formatting.

        Args:
            filename: Path to Word document
            text: Heading text
            level: Heading level (1-9)
            font_name: Font family (e.g., 'Helvetica')
            font_size: Font size in points (e.g., 14)
            bold: Make heading bold
            italic: Make heading italic
            border_bottom: Add bottom border (for section headers)
        """
        return content_tools.add_heading(filename, text, level, font_name, font_size, bold, italic, border_bottom)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Picture",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_picture(filename: str, image_path: str, width: float = None):
        """Add an image to a Word document."""
        return content_tools.add_picture(filename, image_path, width)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Table",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_table(filename: str, rows: int, cols: int, data: list[list[str]] = None):
        """Add a table to a Word document.

        [cross-platform mode] Requires the .docx file to be CLOSED in Word.
        If the file is open in Word, use word_live_add_table instead.
        NOTE: data must be a list of lists (2D array), not a list of strings."""
        return content_tools.add_table(filename, rows, cols, data)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Page Break",
            readOnlyHint=False,
            destructiveHint=False,
        ),
    )
    def add_page_break(filename: str):
        """Add a page break to the document."""
        return content_tools.add_page_break(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Delete Paragraph",
            destructiveHint=True,
        ),
    )
    def delete_paragraph(filename: str, paragraph_index: int):
        """Delete a paragraph from a document."""
        return content_tools.delete_paragraph(filename, paragraph_index)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Search and Replace",
            destructiveHint=True,
        ),
    )
    def search_and_replace(filename: str, find_text: str, replace_text: str):
        """Search for text and replace all occurrences.

        [cross-platform mode] Requires the .docx file to be CLOSED in Word.
        If the file is open in Word, use word_live_replace_text instead."""
        return content_tools.search_and_replace(filename, find_text, replace_text)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Replace Block Below Header",
            readOnlyHint=False,
            destructiveHint=True,
        ),
    )
    def replace_paragraph_block_below_header(filename: str, header_text: str, new_paragraphs: list[str], detect_block_end_fn: str = None):
        """Reemplaza el bloque de párrafos debajo de un encabezado, evitando modificar TOC."""
        return replace_paragraph_block_below_header_tool(filename, header_text, new_paragraphs, detect_block_end_fn)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Replace Block Between Anchors",
            readOnlyHint=False,
            destructiveHint=True,
        ),
    )
    def replace_block_between_manual_anchors(filename: str, start_anchor_text: str, new_paragraphs: list[str], end_anchor_text: str = None, match_fn: str = None, new_paragraph_style: str = None):
        """Replace all content between start_anchor_text and end_anchor_text (or next logical header if not provided)."""
        return replace_block_between_manual_anchors_tool(filename, start_anchor_text, new_paragraphs, end_anchor_text, match_fn, new_paragraph_style)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Table of Contents",
            destructiveHint=True,
        ),
    )
    def add_table_of_contents(filename: str, title: str = "Table of Contents", max_level: int = 3):
        """Add a table of contents based on heading styles."""
        return content_tools.add_table_of_contents(filename, title, max_level)
