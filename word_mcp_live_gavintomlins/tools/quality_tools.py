"""Quality tools: validation, markdown read-back, batch creation (ADR 0003/0004)."""

import json
import os

from docx import Document
from docx.oxml import OxmlElement

from word_mcp_live_gavintomlins.core.markdown_read import document_to_markdown
from word_mcp_live_gavintomlins.core.markdown_write import markdown_to_document
from word_mcp_live_gavintomlins.core.validation import validate_docx
from word_mcp_live_gavintomlins.utils.file_utils import (
    check_file_writeable,
    ensure_docx_extension,
)


async def validate_document(filename: str) -> str:
    """Run business-rule validation and return a structured JSON report."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    try:
        report = validate_docx(filename)
    except Exception as e:
        return f"Failed to validate document: {e}"
    return json.dumps(report, indent=2, ensure_ascii=False)


async def get_document_markdown(filename: str, show_revisions: bool = False) -> str:
    """Render the document as Markdown for content verification."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    try:
        return document_to_markdown(filename, show_revisions=show_revisions)
    except Exception as e:
        return f"Failed to render document as markdown: {e}"


async def create_document_from_markdown(
    filename: str,
    markdown: str,
    title: str = None,
    author: str = None,
    template: str = None,
) -> str:
    """Create a complete document from Markdown in a single call.

    With ``template``, the new document inherits that .docx's styles,
    headers/footers and page setup (the template's body content is not
    carried over).
    """
    filename = ensure_docx_extension(filename)
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot create document: {error_message}"
    if template:
        template = ensure_docx_extension(template)
        if not os.path.exists(template):
            return f"Template {template} does not exist"
    try:
        stats = markdown_to_document(
            markdown, filename, title=title, author=author, template=template
        )
    except Exception as e:
        return f"Failed to create document from markdown: {e}"
    detail = ", ".join(f"{v} {k}" for k, v in stats.items() if v)
    source = f" from template {template}" if template else ""
    return (
        f"Document {filename} created successfully{source} "
        f"({detail or 'empty document'}). "
        "Verify with validate_document and get_document_markdown before delivery."
    )


async def set_update_fields_on_open(filename: str) -> str:
    """Mark the document so Word refreshes fields (TOC, page numbers) on open."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}"
    try:
        doc = Document(filename)
        settings = doc.settings.element
        w_ns = settings.nsmap.get("w") or (
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        )
        tag = f"{{{w_ns}}}updateFields"
        existing = settings.find(tag)
        if existing is None:
            update = OxmlElement("w:updateFields")
            update.set(f"{{{w_ns}}}val", "true")
            settings.insert(0, update)
        else:
            existing.set(f"{{{w_ns}}}val", "true")
        doc.save(filename)
        return (
            f"Document {filename} will now refresh fields (TOC, page numbers) "
            "when opened in Word"
        )
    except Exception as e:
        return f"Failed to set update-fields flag: {e}"
