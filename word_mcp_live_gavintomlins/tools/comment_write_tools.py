"""
Comment writing tools for Word Document Server.

These tools provide MCP interfaces for adding comments to Word documents.
"""

import json
import os

from word_mcp_live_gavintomlins.defaults import DEFAULT_AUTHOR, DEFAULT_INITIALS
from word_mcp_live_gavintomlins.utils.file_utils import check_file_writeable, ensure_docx_extension, get_file_lock
from word_mcp_live_gavintomlins.core.comment_writer import add_comment_to_doc
from word_mcp_live_gavintomlins.core import comment_threads


def _writeable_or_error(filename: str) -> str | None:
    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify document: {error_message}"})
    return None


async def add_comment(
    filename: str,
    target_text: str,
    comment_text: str,
    author: str = DEFAULT_AUTHOR,
    initials: str = DEFAULT_INITIALS,
) -> str:
    """Add a comment to a Word document anchored to specific text.

    Args:
        filename: Path to Word document
        target_text: Text in the document to attach the comment to
        comment_text: The comment content
        author: Comment author name
        initials: Author initials

    Returns:
        JSON string with result
    """
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify document: {error_message}"})

    if not target_text:
        return json.dumps({"success": False, "error": "target_text cannot be empty"})
    if not comment_text:
        return json.dumps({"success": False, "error": "comment_text cannot be empty"})

    try:
        async with get_file_lock(filename):
            result = add_comment_to_doc(filename, target_text, comment_text, author, initials)
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to add comment: {str(e)}"})


async def reply_to_comment(
    filename: str,
    comment_id: int,
    reply_text: str,
    author: str = DEFAULT_AUTHOR,
    initials: str = DEFAULT_INITIALS,
) -> str:
    """Add a threaded reply to an existing comment (file must be closed)."""
    filename = ensure_docx_extension(filename)
    error = _writeable_or_error(filename)
    if error:
        return error
    if not reply_text:
        return json.dumps({"success": False, "error": "reply_text cannot be empty"})
    try:
        async with get_file_lock(filename):
            result = comment_threads.reply_to_comment(
                filename, comment_id, reply_text, author, initials
            )
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to add reply: {str(e)}"})


async def resolve_comment(
    filename: str,
    comment_id: int,
    resolved: bool = True,
) -> str:
    """Mark a comment thread resolved or reopen it (file must be closed)."""
    filename = ensure_docx_extension(filename)
    error = _writeable_or_error(filename)
    if error:
        return error
    try:
        async with get_file_lock(filename):
            result = comment_threads.resolve_comment(filename, comment_id, resolved)
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to resolve comment: {str(e)}"})


async def delete_comment(
    filename: str,
    comment_id: int,
) -> str:
    """Delete a comment and its replies (file must be closed)."""
    filename = ensure_docx_extension(filename)
    error = _writeable_or_error(filename)
    if error:
        return error
    try:
        async with get_file_lock(filename):
            result = comment_threads.delete_comment(filename, comment_id)
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to delete comment: {str(e)}"})
