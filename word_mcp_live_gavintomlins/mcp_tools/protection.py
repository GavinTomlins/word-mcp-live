"""Protection tools: encryption, restricted editing, signatures, verification.

Registrations moved verbatim from ``main.py``'s former ``register_tools()``.
"""

from mcp.types import ToolAnnotations

from word_mcp_live_gavintomlins.tools import protection_tools


def register(mcp):
    """Register protection tools on ``mcp``."""
    # Protection tools
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Protect Document",
            readOnlyHint=False,
            destructiveHint=True,
        ),
    )
    def protect_document(filename: str, password: str, confirm_password: str = None):
        """Add password protection (encryption) to a Word document.

        The password is NOT stored anywhere — it is discarded after
        encryption.  Make sure to remember it.

        Include confirm_password for extra safety: the LLM asks the user
        to type the password twice and passes both; they must match.
        """
        return protection_tools.protect_document(filename, password, confirm_password)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Unprotect Document",
            readOnlyHint=False,
            destructiveHint=True,
        ),
    )
    def unprotect_document(filename: str, password: str):
        """Remove password protection from a Word document."""
        return protection_tools.unprotect_document(filename, password)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Restricted Editing",
            destructiveHint=True,
        ),
    )
    def add_restricted_editing(filename: str, password: str, editable_sections: list[str]):
        """Add restricted editing to a document, allowing editing only in specified sections."""
        return protection_tools.add_restricted_editing(filename, password, editable_sections)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Signature Block",
            destructiveHint=True,
        ),
    )
    def add_signature_block(filename: str, signer_name: str,
                            title: str = None, reason: str = None,
                            location: str = None, show_date: bool = True):
        """Add a visible signature block to a Word document.

        Cross-platform: inserts a formal text signature block (separator,
        signer name, title, date, location, signature line).

        Windows + COM: additionally creates a native Word Signature Line
        shape that the recipient can double-click to sign.

        NOTE: This is a signature placeholder, not a cryptographic
        PKI/X.509 digital signature.
        """
        return protection_tools.add_signature_block(
            filename, signer_name,
            title=title, reason=reason,
            location=location, show_date=show_date,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Verify Document",
            readOnlyHint=True,
        ),
    )
    def verify_document(filename: str, password: str = None):
        """Verify document protection status (restricted editing, encryption, signature lines)."""
        return protection_tools.verify_document(filename, password)
