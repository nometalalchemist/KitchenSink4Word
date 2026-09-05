"""KitchenSink4Word: a .docx MCP server engineered not to corrupt."""

#: The running build. Kept in step with pyproject.toml [project] version by
#: tests/unit/test_update_check.py, so a release bump cannot land in one
#: place only. The update check reports this number, not installed metadata
#: (an editable install's metadata goes stale the moment a version is
#: bumped, and a wrong "you are running X" line is worse than none).
__version__ = "2.0.2"
