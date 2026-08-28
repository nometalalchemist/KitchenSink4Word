"""Workflow guidance: recommended tool sequences for common multi-step tasks.

With 175+ tools, discoverability is a real problem for both agents and
integrators (API review, item 7). This module holds the curated sequences as
plain data; get_workflows serves them. Every tool named in a step MUST be a
registered tool on the server (tests assert this against the live registry),
so a workflow can never point at a tool that does not exist.

The sequences are recommendations, not scripts: steps marked optional can be
skipped, and the notes carry the judgment calls a caller should know about.
"""

from __future__ import annotations

from ..core.errors import WordMcpError

# task -> {summary, steps: [{tool, why, optional?}], notes: [...]}
WORKFLOWS: dict[str, dict] = {
    "process-feedback": {
        "summary": (
            "Work through advisor/reviewer feedback: read tracked changes "
            "and comments, act on them, verify nothing broke."
        ),
        "steps": [
            {"tool": "get_tracked_changes",
             "why": "see every insertion/deletion with author, date, and text"},
            {"tool": "revision_summary",
             "why": "counts by author and type before deciding what to accept"},
            {"tool": "comment_report",
             "why": "all comments with their anchored text, threading, and resolved state"},
            {"tool": "accept_revisions",
             "why": "accept changes (filter by author/type); reject_revisions is the mirror"},
            {"tool": "resolve_comment",
             "why": "mark handled comments resolved so the remaining work stays visible"},
            {"tool": "validate_document",
             "why": "confirm note and field structure survived the revision pass"},
        ],
        "notes": [
            "Work on a fresh copy (see the heavy-editing workflow) when the "
            "original must stay untouched for comparison.",
            "structured_diff compares the before/after files if you kept both.",
        ],
    },
    "prepare-submission": {
        "summary": (
            "Pre-submission quality gate: snapshot, validate structure and "
            "citations, check accessibility, then clean for submission."
        ),
        "steps": [
            {"tool": "create_snapshot",
             "why": "permanent DTG-stamped keeper before any submission surgery"},
            {"tool": "validate_document",
             "why": "package integrity, note consistency, balanced fields"},
            {"tool": "check_citation_parity",
             "why": "in-text citations vs reference list, both directions"},
            {"tool": "check_reference_field_integrity",
             "why": "reference-manager fields (Zotero/Mendeley/EndNote) still intact"},
            {"tool": "validate_cross_references",
             "why": "no broken REF/PAGEREF targets"},
            {"tool": "word_count_with_exclusions",
             "why": "journal-style count (excluding references, captions, etc.)"},
            {"tool": "prepare_for_submission",
             "why": "the cleanup pass itself: comments, revisions, metadata"},
            {"tool": "anonymize_for_review", "optional": True,
             "why": "blind-review venues only; deanonymize_document reverses it"},
        ],
        "notes": [
            "audit_accessibility is worth adding for publishers that check it.",
            "Run the validation steps again AFTER prepare_for_submission; the "
            "cleanup itself is a mutation.",
        ],
    },
    "format-citations": {
        "summary": (
            "Citation work without breaking the reference manager: detect "
            "which system owns the citations first, then act within it."
        ),
        "steps": [
            {"tool": "detect_citation_system",
             "why": "learn which system (Word native, Zotero, Mendeley, EndNote, "
                    "plain text) owns the citations BEFORE touching them; mixing "
                    "systems creates a split-brain bibliography"},
            {"tool": "list_reference_fields",
             "why": "inventory every manager field with location and cached text"},
            {"tool": "check_citation_parity",
             "why": "find cited-but-not-listed and listed-but-never-cited entries"},
            {"tool": "convert_citation_style", "optional": True,
             "why": "plain-text style conversion (APA/Chicago/etc.) when the "
                    "document is NOT manager-managed"},
            {"tool": "check_reference_field_integrity",
             "why": "post-edit check that no field pair was orphaned"},
        ],
        "notes": [
            "If detect_citation_system reports split_brain=true, stop and "
            "resolve it with the user before adding any citation.",
            "insert_citation / add_source are Word-native; insert_zotero_citation "
            "is Zotero. Use the family the document already uses.",
        ],
    },
    "build-from-template": {
        "summary": (
            "Produce a document from an institutional template without "
            "contaminating the template itself."
        ),
        "steps": [
            {"tool": "copy_document",
             "why": "never build inside the template file; work on a copy"},
            {"tool": "list_template_placeholders",
             "why": "see every {{placeholder}} the template expects"},
            {"tool": "fill_template",
             "why": "fill the placeholders from a values map"},
            {"tool": "apply_template", "optional": True,
             "why": "import styles/formatting from a reference document when the "
                    "content came from elsewhere"},
            {"tool": "check_template_compliance",
             "why": "verify the result still matches the template's rules"},
            {"tool": "validate_document",
             "why": "structural check before handing the file over"},
        ],
        "notes": [
            "Templates using content controls instead of {{placeholders}}: use "
            "list_form_fields + fill_form_fields.",
        ],
    },
    "heavy-editing": {
        "summary": (
            "Sustained editing sessions (tester-recommended pattern): the "
            "user keeps the document open in Word for reading while the "
            "agent works on a fresh DTG-stamped copy with full file-mode "
            "tool access, no COM round-trips, no lock conflicts."
        ),
        "steps": [
            {"tool": "copy_document",
             "why": "make the fresh DTG-named working copy (YYYYMMDD_HHMM_Name."
                    "docx); create_snapshot automates the DTG naming"},
            {"tool": "get_outline",
             "why": "orient on the document's structure before editing"},
            {"tool": "get_text",
             "why": "read the sections being edited (slice with start/end)"},
            {"tool": "search_and_replace",
             "why": "bulk text edits on the copy at full file-mode speed"},
            {"tool": "batch_apply", "optional": True,
             "why": "group many small edits into one atomic pass"},
            {"tool": "validate_document",
             "why": "structural check after the editing pass"},
            {"tool": "word_count",
             "why": "sanity-check the result against expectations"},
        ],
        "notes": [
            "The original stays open in Word untouched; the user opens the new "
            "file when the pass is done. This was the most productive pattern "
            "observed in real dissertation work.",
            "Every mutation on the copy still auto-backs up to .ks4w-backups/, "
            "so even the working copy has prev/anchor rollback.",
        ],
    },
}


def get_workflows(task: str | None = None) -> dict:
    """No task: list available tasks with summaries. With a task: the
    recommended tool sequence, one why-line per step, plus notes."""
    if task is None:
        return {
            "tasks": [
                {"task": name, "summary": wf["summary"]}
                for name, wf in WORKFLOWS.items()
            ],
            "note": "call again with task='<name>' for the step-by-step sequence",
        }
    wf = WORKFLOWS.get(task)
    if wf is None:
        raise WordMcpError(
            f"unknown task {task!r}; available tasks: {sorted(WORKFLOWS)}"
        )
    return {"task": task, **wf}
