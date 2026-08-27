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


class WordNotRunning(WordMcpError):
    """No attachable interactive Word instance (live tools need one)."""


class DocumentNotOpenInWord(WordMcpError):
    """Live tool targeted a document that is not open in the running Word."""


class ProtectedViewRefused(WordMcpError):
    """Document is in Protected View; the user must click Enable Editing."""


class WordBusy(WordMcpError):
    """Word rejected the call (dialog, Backstage, or a running command)."""


class WordBlocked(WordMcpError):
    """Word is not answering at all (long synchronous operation in progress)."""


class WordDisconnected(WordMcpError):
    """Word or the document closed mid-call; the edit may be partially applied."""
