"""Quality tools: validation, markdown read-back, batch creation."""

from mcp.types import ToolAnnotations

from word_mcp_live_gavintomlins.tools import quality_tools


def register(mcp):
    """Register quality tools on ``mcp``."""

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Validate Document",
            readOnlyHint=True,
        ),
    )
    def validate_document(filename: str):
        """Run business-rule validation on a Word document and return a JSON report.

        Checks table grid consistency (skewed column widths), image aspect
        ratio distortion, comment integrity across the comment part files,
        stale TOC fields, missing xml:space on whitespace, and OOXML element
        order. Run this after creating or editing a document, before
        reporting the work complete.

        [cross-platform mode] Requires the .docx file to be CLOSED in Word.
        """
        return quality_tools.validate_document(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get Document Markdown",
            readOnlyHint=True,
        ),
    )
    def get_document_markdown(filename: str, show_revisions: bool = False):
        """Render a Word document as Markdown to verify structure and content.

        Preserves headings, lists, tables, bold/italic and hyperlinks. With
        show_revisions=True, tracked insertions render as {++text++} and
        deletions as {--text--}. Use this to confirm a document matches what
        you intended to write.

        [cross-platform mode] Requires the .docx file to be CLOSED in Word.
        If the file is open in Word, use word_live_get_text instead.
        """
        return quality_tools.get_document_markdown(filename, show_revisions)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Create Document from Markdown",
            destructiveHint=True,
        ),
    )
    def create_document_from_markdown(
        filename: str,
        markdown: str,
        title: str = None,
        author: str = None,
        template: str = None,
    ):
        """Create a complete Word document from Markdown in one call.

        Supports #–###### headings, paragraphs, **bold**, *italic*, `code`,
        [links](url), nested bullet/numbered lists (2-space indent per
        level), pipe tables (header row bold), and --- horizontal rules.

        Pass template=<path to a .docx> to build on that template: the new
        document inherits its styles (headings, fonts, colors),
        headers/footers and page setup, while the template's own body
        content is not carried over. Use this for branded deliverables.

        Prefer this over many add_paragraph/add_heading calls when creating
        a new document. After creation, verify with validate_document and
        get_document_markdown.
        """
        return quality_tools.create_document_from_markdown(
            filename, markdown, title, author, template
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Set Update Fields on Open",
            destructiveHint=False,
        ),
    )
    def set_update_fields_on_open(filename: str):
        """Make Word refresh fields (TOC, page numbers) when the document opens.

        Use after adding a table of contents or page-number fields so the
        reader is not shown stale values.

        [cross-platform mode] Requires the .docx file to be CLOSED in Word.
        """
        return quality_tools.set_update_fields_on_open(filename)
