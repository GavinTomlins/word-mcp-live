"""Typed runtime configuration for the Word Document MCP Server.

All settings come from environment variables or a ``.env`` file. New-style
names use the ``WORD_MCP_`` prefix; the names accepted by earlier releases
(``MCP_TRANSPORT``, ``MCP_HOST``, ``PORT``/``MCP_PORT``, ``MCP_PATH``,
``WORD_MCP_LIVE_API_KEY``, ``WORD_MCP_LIVE_INSECURE``, ``MCP_AUTHOR``,
``MCP_AUTHOR_INITIALS``) keep working as aliases.
"""

from pydantic import AliasChoices, Field, field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Validated server configuration."""

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    transport: str = Field(
        "stdio",
        validation_alias=AliasChoices("WORD_MCP_TRANSPORT", "MCP_TRANSPORT"),
        description="'stdio' (default) or 'http' (streamable HTTP).",
    )
    host: str = Field(
        "127.0.0.1",
        validation_alias=AliasChoices("WORD_MCP_HOST", "MCP_HOST"),
    )
    port: int = Field(
        8000,
        # PORT comes first after the new name so Render's injected PORT
        # keeps taking precedence over MCP_PORT, as before.
        validation_alias=AliasChoices("WORD_MCP_PORT", "PORT", "MCP_PORT"),
    )
    path: str = Field(
        "/mcp",
        validation_alias=AliasChoices("WORD_MCP_PATH", "MCP_PATH"),
    )
    api_key: str | None = Field(
        None,
        validation_alias=AliasChoices("WORD_MCP_API_KEY", "WORD_MCP_LIVE_API_KEY"),
        description="Bearer token required for HTTP transport.",
    )
    insecure: bool = Field(
        False,
        validation_alias=AliasChoices("WORD_MCP_INSECURE", "WORD_MCP_LIVE_INSECURE"),
        description="Explicit opt-in to run HTTP without authentication (dev only).",
    )
    log_level: str = Field(
        "INFO",
        validation_alias=AliasChoices("WORD_MCP_LOG_LEVEL", "FASTMCP_LOG_LEVEL"),
    )
    mask_errors: bool | None = Field(
        None,
        validation_alias=AliasChoices("WORD_MCP_MASK_ERRORS"),
        description="Override error masking; defaults to on for HTTP, off for stdio.",
    )
    author: str = Field(
        "Author",
        validation_alias=AliasChoices("WORD_MCP_AUTHOR", "MCP_AUTHOR"),
        description="Default author for comments and tracked changes.",
    )
    author_initials: str = Field(
        "",
        validation_alias=AliasChoices("WORD_MCP_AUTHOR_INITIALS", "MCP_AUTHOR_INITIALS"),
    )

    @field_validator("transport")
    @classmethod
    def _normalize_transport(cls, value: str) -> str:
        value = value.lower()
        if value == "streamable-http":  # legacy spelling
            return "http"
        if value == "sse":
            raise ValueError(
                "The SSE transport was removed (deprecated in the MCP spec). "
                "Use WORD_MCP_TRANSPORT=http instead."
            )
        if value not in ("stdio", "http"):
            raise ValueError(
                f"Invalid transport {value!r}: expected 'stdio' or 'http'."
            )
        return value

    @property
    def mask_error_details(self) -> bool:
        """Mask exception details from clients unless explicitly overridden.

        Defaults to masking on HTTP (internet-exposed deployments must not
        leak tracebacks or filesystem paths) and full details on stdio.
        """
        if self.mask_errors is not None:
            return self.mask_errors
        return self.transport != "stdio"

    @property
    def auth_required(self) -> bool:
        """True when HTTP is requested but neither a key nor the insecure opt-in is set."""
        return self.transport != "stdio" and not self.api_key and not self.insecure


def get_settings() -> Settings:
    """Read settings from the environment / .env file."""
    return Settings()
