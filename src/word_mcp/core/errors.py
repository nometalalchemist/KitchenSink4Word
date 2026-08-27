"""Exception types for word-mcp. Every tool call maps these to actionable messages."""


class WordMcpError(Exception):
    """Base class; message text is user-facing."""


class DocumentNotFound(WordMcpError):
    pass


class DocumentLocked(WordMcpError):
    """File is open in Word (or another process holds a lock)."""


class DocumentCorrupt(WordMcpError):
    """File is not a valid .docx package."""


class DocumentProtected(WordMcpError):
    """File is encrypted/password-protected."""


class TargetNotFound(WordMcpError):
    """Anchor text, paragraph, table, or note the tool was told to act on does not exist."""


class AmbiguousTarget(WordMcpError):
    """More than one match for the given anchor; caller must disambiguate."""


class UnsupportedStructure(WordMcpError):
    """Merge topology or XML shape we refuse to guess about (conservative mode)."""


class ValidationFailed(WordMcpError):
    """Post-edit validation caught a problem; the original file was NOT modified."""
