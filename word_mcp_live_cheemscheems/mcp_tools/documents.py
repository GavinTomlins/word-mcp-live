"""Document-level tools: create, copy, inspect, backup, merge.

Registrations moved verbatim from ``main.py``'s former ``register_tools()``.
"""

from mcp.types import ToolAnnotations

from word_mcp_live_cheemscheems.tools import document_tools


def register(mcp):
    """Register documents tools on ``mcp``."""
    # Document tools (create, copy, info, etc.)
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Create Word Document",
            destructiveHint=True,
        ),
    )
    def create_document(filename: str, title: str = None, author: str = None):
        """Create a new Word document with optional metadata."""
        return document_tools.create_document(filename, title, author)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Copy Word Document",
            destructiveHint=True,
        ),
    )
    def copy_document(source_filename: str, destination_filename: str = None):
        """Create a copy of a Word document."""
        return document_tools.copy_document(source_filename, destination_filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get Document Info",
            readOnlyHint=True,
        ),
    )
    def get_document_info(filename: str):
        """Get information about a Word document."""
        return document_tools.get_document_info(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get Document Text",
            readOnlyHint=True,
        ),
    )
    def get_document_text(filename: str, show_revisions: bool = False):
        """Extract all text from a Word document.

        [cross-platform mode] Requires the .docx file to be CLOSED in Word.
        If the file is open in Word, use word_live_get_text instead.

        By default returns the effective final text (insertions applied,
        deletions removed).  Set show_revisions=True to get inline redline
        markup where deletions appear as [-deleted-] and insertions as
        {+inserted+}."""
        return document_tools.get_document_text(filename, show_revisions=show_revisions)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get Document Outline",
            readOnlyHint=True,
        ),
    )
    def get_document_outline(filename: str):
        """Get the structure of a Word document."""
        return document_tools.get_document_outline(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="List Available Documents",
            readOnlyHint=True,
        ),
    )
    def list_available_documents(directory: str = "."):
        """List all .docx files in the specified directory."""
        return document_tools.list_available_documents(directory)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get Document XML",
            readOnlyHint=True,
        ),
    )
    def get_document_xml(filename: str):
        """Get the raw XML structure of a Word document."""
        return document_tools.get_document_xml_tool(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Backup Document",
            readOnlyHint=True,
        ),
    )
    def backup_document(filename: str, note: str = None):
        """Create a backup copy of the document.

        Use this before making large or destructive changes (encryption,
        large-scale replace, etc.).  Backups are stored in a ``_backup``
        folder next to the source file.
        """
        return document_tools.backup_document(filename, note)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Merge Documents",
            destructiveHint=True,
        ),
    )
    def merge_documents(target_filename: str, source_filenames: list[str], add_page_breaks: bool = True):
        """Merge multiple Word documents into a single target document."""
        return document_tools.merge_documents(target_filename, source_filenames, add_page_breaks)
