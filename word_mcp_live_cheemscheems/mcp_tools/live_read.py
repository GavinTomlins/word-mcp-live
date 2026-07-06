"""Live read tools (Word must be running with the document open).

Registrations moved verbatim from ``main.py``'s former ``register_tools()``.
"""

from mcp.types import ToolAnnotations

from word_mcp_live_cheemscheems.defaults import DEFAULT_AUTHOR
from word_mcp_live_cheemscheems.tools import live_read_tools

LIVE_TAGS = {"live"}


def register(mcp):
    """Register live read tools on ``mcp``."""
    # --- Live read tools (Windows only, requires Word running) ---

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Get Text",
            readOnlyHint=True,
        ),
    )
    def word_live_get_text(filename: str = None):
        """[Windows only] Get all text from a Word document open in Word, paragraph by paragraph. For large documents (200+ paragraphs), automatically returns only the first 3 pages — use word_live_get_page_text to read specific pages."""
        return live_read_tools.word_live_get_text(filename)

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Take Snapshot",
            readOnlyHint=True,
        ),
    )
    def word_live_take_snapshot(filename: str = None):
        """[Windows only] Store a snapshot of the current document text for later diffing without returning the full text. Use word_live_get_diff afterwards to see what changed."""
        return live_read_tools.word_live_take_snapshot(filename)

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Get Diff",
            readOnlyHint=True,
        ),
    )
    def word_live_get_diff(filename: str = None):
        """[Windows only] Return only paragraphs that changed since the last snapshot. Compares current document against snapshot from word_live_take_snapshot. Returns added, modified, deleted paragraphs. Automatically updates snapshot after diffing."""
        return live_read_tools.word_live_get_diff(filename)

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Snapshot Status",
            readOnlyHint=True,
        ),
    )
    def word_live_snapshot_status(filename: str = None):
        """[Windows only] Check whether a snapshot exists for the document and how old it is. Returns has_snapshot, age_seconds, and paragraph_count."""
        return live_read_tools.word_live_snapshot_status(filename)

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Get Paragraph Format",
            readOnlyHint=True,
        ),
        description=live_read_tools.word_live_get_paragraph_format.__doc__,
    )
    def word_live_get_paragraph_format(
        filename: str = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        include_runs: bool = False,
    ):
        return live_read_tools.word_live_get_paragraph_format(
            filename, start_paragraph, end_paragraph, include_runs,
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Get Info",
            readOnlyHint=True,
        ),
    )
    def word_live_get_info(filename: str = None):
        """[Windows only] Get document info (pages, words, sections, etc.) from a Word document open in Word. Requires Word running."""
        return live_read_tools.word_live_get_info(filename)

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Set Core Properties",
            destructiveHint=True,
        ),
        description=live_read_tools.word_live_set_core_properties.__doc__,
    )
    def word_live_set_core_properties(
        filename: str = None,
        title: str = None,
        subject: str = None,
        author: str = None,
        keywords: str = None,
        comments: str = None,
        category: str = None,
        manager: str = None,
        company: str = None,
        last_author: str = None,
    ):
        return live_read_tools.word_live_set_core_properties(
            filename=filename,
            title=title,
            subject=subject,
            author=author,
            keywords=keywords,
            comments=comments,
            category=category,
            manager=manager,
            company=company,
            last_author=last_author,
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live List Open",
            readOnlyHint=True,
        ),
    )
    def word_live_list_open():
        """[Windows only] List all documents currently open in Word with name, path, pages, and saved status."""
        return live_read_tools.word_live_list_open()

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Find Text",
            readOnlyHint=True,
        ),
    )
    def word_live_find_text(
        filename: str = None,
        search_text: str = "",
        match_case: bool = False,
        whole_word: bool = False,
        use_wildcards: bool = False,
        context_chars: int = 60,
        max_results: int = 50,
    ):
        """[Windows only] Find text in a Word document open in Word. Returns positions and context.
        With use_wildcards=True, supports ^m (page break), ^t (tab), ^p (paragraph mark) and Word wildcards.
        context_chars controls how many characters of surrounding context to return (default 60). Requires Word running."""
        return live_read_tools.word_live_find_text(
            filename, search_text, match_case, whole_word,
            use_wildcards, context_chars, max_results,
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Get Comments",
            readOnlyHint=True,
        ),
    )
    def word_live_get_comments(filename: str = None):
        """[Windows only] Get all comments from a Word document open in Word. Requires Word running."""
        return live_read_tools.word_live_get_comments(filename)

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Add Comment",
            destructiveHint=True,
        ),
    )
    def word_live_add_comment(
        filename: str = None,
        start: int = None,
        end: int = None,
        paragraph_index: int = None,
        text: str = "",
        author: str = DEFAULT_AUTHOR,
    ):
        """[Windows only] Add a comment to a Word document open in Word.
        Specify start/end character positions or paragraph_index (1-indexed). Requires Word running."""
        return live_read_tools.word_live_add_comment(
            filename, start, end, paragraph_index, text, author
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Reply to Comment",
            destructiveHint=True,
        ),
    )
    def word_live_reply_to_comment(
        filename: str = None,
        comment_index: int = None,
        text: str = "",
        author: str = DEFAULT_AUTHOR,
    ):
        """[Windows only] Reply to an existing comment in a Word document open in Word.
        Adds a threaded reply. Use word_live_get_comments to find the comment_index.
        Requires Word 2016+ running."""
        return live_read_tools.word_live_reply_to_comment(
            filename, comment_index, text, author
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Resolve Comment",
            destructiveHint=True,
        ),
    )
    def word_live_resolve_comment(
        filename: str = None,
        comment_index: int = None,
        resolve: bool = True,
    ):
        """[Windows only] Resolve or unresolve a comment in a Word document open in Word.
        Sets the comment's Done property. Use word_live_get_comments to find comment_index.
        Requires Word 2016+ running."""
        return live_read_tools.word_live_resolve_comment(
            filename, comment_index, resolve
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Delete Comment",
            destructiveHint=True,
        ),
    )
    def word_live_delete_comment(
        filename: str = None,
        comment_index: int = None,
    ):
        """[Windows only] Delete a comment from a Word document open in Word.
        Permanently removes the comment. Use word_live_get_comments to find comment_index.
        Requires Word running."""
        return live_read_tools.word_live_delete_comment(
            filename, comment_index
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live List Revisions",
            readOnlyHint=True,
        ),
    )
    def word_live_list_revisions(filename: str = None):
        """[Windows only] List all tracked changes (revisions) in a Word document open in Word. Requires Word running."""
        return live_read_tools.word_live_list_revisions(filename)

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Accept Revisions",
            destructiveHint=True,
        ),
    )
    def word_live_accept_revisions(
        filename: str = None,
        author: str = None,
        revision_ids: list[int] = None,
    ):
        """[Windows only] Accept tracked changes in a Word document open in Word.
        Filter by author or specific revision IDs. Requires Word running."""
        return live_read_tools.word_live_accept_revisions(
            filename, author, revision_ids
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Reject Revisions",
            destructiveHint=True,
        ),
    )
    def word_live_reject_revisions(
        filename: str = None,
        author: str = None,
        revision_ids: list[int] = None,
    ):
        """[Windows only] Reject tracked changes in a Word document open in Word.
        Filter by author or specific revision IDs. Requires Word running."""
        return live_read_tools.word_live_reject_revisions(
            filename, author, revision_ids
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Get Page Text",
            readOnlyHint=True,
        ),
        description=live_read_tools.word_live_get_page_text.__doc__,
    )
    def word_live_get_page_text(
        filename: str = None,
        page: int = 1,
        end_page: int = None,
    ):
        return live_read_tools.word_live_get_page_text(
            filename, page, end_page,
        )

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Get Undo History",
            readOnlyHint=True,
        ),
    )
    def word_live_get_undo_history(filename: str = None):
        """[Windows only] Get the undo stack from a Word document open in Word.
        Shows MCP tool operations as named entries. Requires Word running."""
        return live_read_tools.word_live_get_undo_history(filename)

    @mcp.tool(
        tags=LIVE_TAGS,
        annotations=ToolAnnotations(
            title="Word Live Diagnose Layout",
            readOnlyHint=True,
        ),
        description=live_read_tools.word_live_diagnose_layout.__doc__,
    )
    def word_live_diagnose_layout(filename: str = None):
        return live_read_tools.word_live_diagnose_layout(filename)
