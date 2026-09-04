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

# Tools named in steps that are DECLARED but not yet registered. Empty
# since the Phase 3 view/batch layer landed (get_document_view and
# apply_edits are registered); the registry test (test_absorptions.py)
# treats entries here as pending rather than missing.
PENDING_TOOLS: set[str] = set()

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
            {"tool": "get_revision_report",
             "why": "counts by author and type (mode='summary') before "
                    "deciding what to accept"},
            {"tool": "comment_report",
             "why": "all comments with their anchored text, threading, and resolved state"},
            {"tool": "resolve_revisions",
             "why": "action='accept' applies changes (filter by author); "
                    "action='reject' is the mirror"},
            {"tool": "manage_comment",
             "why": "action='resolve' marks handled comments so the "
                    "remaining work stays visible"},
            {"tool": "validate",
             "why": "confirm note and field structure survived the revision "
                    "pass (default checks=['core'])"},
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
            {"tool": "manage_backups",
             "why": "action='snapshot': permanent DTG-stamped keeper before "
                    "any submission surgery"},
            {"tool": "validate",
             "why": "one battery: checks=['core', 'citation_parity', "
                    "'reference_fields', 'cross_references', "
                    "'accessibility']"},
            {"tool": "word_count",
             "why": "journal-style count via exclusions=['references', "
                    "'captions', ...]"},
            {"tool": "prepare_for_submission",
             "why": "the cleanup pass itself: comments, revisions, metadata"},
            {"tool": "anonymize_for_review", "optional": True,
             "why": "blind-review venues only; deanonymize_document reverses it"},
        ],
        "notes": [
            "Run the validate battery again AFTER prepare_for_submission; the "
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
            {"tool": "list_elements",
             "why": "type='reference_fields' inventories every manager field "
                    "with location and cached text"},
            {"tool": "validate",
             "why": "checks=['citation_parity'] finds cited-but-not-listed "
                    "and listed-but-never-cited entries"},
            {"tool": "convert_citation_style", "optional": True,
             "why": "plain-text style conversion (APA/Chicago/etc.) when the "
                    "document is NOT manager-managed"},
            {"tool": "validate",
             "why": "checks=['reference_fields'] after editing: no field "
                    "pair was orphaned"},
        ],
        "notes": [
            "If detect_citation_system reports split_brain=true, stop and "
            "resolve it with the user before adding any citation.",
            "insert_citation / manage_source are Word-native; "
            "insert_zotero_citation is Zotero. Use the family the document "
            "already uses.",
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
            {"tool": "list_elements",
             "why": "type='template_placeholders': every {{placeholder}} "
                    "the template expects"},
            {"tool": "fill_template",
             "why": "fill the placeholders from a values map"},
            {"tool": "apply_template", "optional": True,
             "why": "import styles/formatting from a reference document when the "
                    "content came from elsewhere"},
            {"tool": "validate",
             "why": "checks=['template'] (with options.template rules) plus "
                    "the default core check before handing the file over"},
        ],
        "notes": [
            "Templates using content controls instead of {{placeholders}}: "
            "use list_elements(type='form_fields') + set_form_fields.",
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
                    "docx); manage_backups action='snapshot' automates the "
                    "DTG naming"},
            {"tool": "get_outline",
             "why": "orient on the document's structure before editing"},
            {"tool": "get_text",
             "why": "read the sections being edited (slice with start/end)"},
            {"tool": "search_and_replace",
             "why": "bulk text edits on the copy at full file-mode speed"},
            {"tool": "validate",
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
    "live-editing": {
        "summary": (
            "Edit a document while it is OPEN in Word (the Option C save "
            "model): live edits land in the window unsaved; anchors and "
            "views read the last SAVED state, so save before any "
            "anchor-addressed step. COM calls serialize server-side (one "
            "at a time), so concurrent agents queue instead of colliding."
        ),
        "steps": [
            {"tool": "com_word_status",
             "why": "confirm Word is responsive, the document is open, and "
                    "no modal dialog or running COM operation blocks it "
                    "(pending_dialogs / com_serialization)"},
            {"tool": "get_text",
             "why": "read the open document live (dual-mode tools "
                    "auto-route; reads reflect the live state)"},
            {"tool": "search_and_replace",
             "why": "make live edits (also insert_paragraphs, "
                    "set_paragraph_text, format_text, set_cells...); each "
                    "call is one Ctrl+Z step, unsaved until saved"},
            {"tool": "com_save_document",
             "why": "REQUIRED before anchor work: live edits stay unsaved, "
                    "and get_document_view/apply_edits anchors resolve "
                    "from the last saved state"},
            {"tool": "get_document_view",
             "why": "take fresh anchors from the just-saved state (on an "
                    "open document the view says live:true + saved-state "
                    "note)"},
            {"tool": "apply_edits",
             "why": "batch anchor-addressed edits; on the open document "
                    "the batch runs live in one undo group with all "
                    "validation before any write"},
            {"tool": "live_scroll_to",
             "why": "optional: show the user what changed without moving "
                    "their cursor"},
        ],
        "notes": [
            "The chain live edit -> stale view -> STALE_ANCHOR is the "
            "documented trap: a view taken before com_save_document does "
            "not see live edits. Save first, then view, then apply_edits.",
            "diagnose_document has no live route (reads saved XML); "
            "com_validate_opens_clean checks the open copy instead.",
            "If a save fails repeatedly or Word stops answering, "
            "com_word_status reports pending dialogs a human must "
            "dismiss; live_repair fixes crashed-client state.",
        ],
    },
    "comment-partner": {
        "summary": (
            "PREVIEW recipe using shipped tools only: collaborate through "
            "Word comments on an open document — poll for comments "
            "addressed to the AI, fix as tracked changes, reply/resolve "
            "after saving. Live in-thread reply/resolve arrives in v2.1."
        ),
        "steps": [
            {"tool": "get_comments",
             "why": "poll the open document live for new unresolved "
                    "comments addressed to the AI (filter by text "
                    "convention, e.g. '@ai' or a name prefix)"},
            {"tool": "get_text",
             "why": "read the anchored_text's surrounding context live "
                    "before acting"},
            {"tool": "search_and_replace",
             "why": "make the requested fix as a TRACKED change "
                    "(track=true, author='AI name') so the human keeps "
                    "accept/reject control; set_paragraph_text and "
                    "insert_paragraphs work the same way"},
            {"tool": "com_save_document",
             "why": "persist the tracked edits; comment management below "
                    "is file-mode and reads the saved state"},
            {"tool": "manage_comment",
             "why": "reply to the comment describing what was done, then "
                    "action='resolve' it (file mode; v2.1 adds live "
                    "in-thread reply/resolve)"},
        ],
        "notes": [
            "Etiquette: act ONLY on comments addressed to the AI; leave "
            "the humans' discussion threads alone.",
            "Every edit goes in tracked, attributed to the AI author — "
            "the human reviews with accept/reject as usual.",
            "An ambiguous comment gets a REPLY asking for clarification, "
            "never a guessed edit.",
            "manage_comment needs the file saved (and briefly closed if "
            "Word holds the lock: com_save_document close=true, then the "
            "user reopens) — this friction is what v2.1's live comment "
            "route removes.",
        ],
    },
    "migrate-from-v1": {
        "summary": (
            "Translate v1.x call sites to the v2 surface using the shipped "
            "migration map (no capability was lost; 189 v1 tools map onto "
            "the consolidated v2 set)."
        ),
        "steps": [],
        "notes": [
            "migration/v1_to_v2.json (shipped in the repo and the package) "
            "has one entry per v1.6 tool: 'to' names the v2 tool, 'inject' "
            "gives literal v2 params reproducing the v1 behavior, 'params' "
            "maps old param names to new paths (dot notation into nested "
            "objects like the location object).",
            "Positioning params (after_index / after_anchor / at_end) became "
            "the v2 location object: omit location for document end, "
            "{'search': {'text': ...}} for anchors, {'paragraph': N} for "
            "indices.",
            "Enumerators (list_*) merged into list_elements(type=...); "
            "validators and checkers merged into validate(checks=[...]); "
            "note/comment/source lifecycles merged into manage_note / "
            "manage_comment / manage_source.",
            "The narrative guide (docs/MIGRATION_V2.md in the repo) "
            "carries the details, including "
            "the insert_document indexing caveat (v1 after_index counted "
            "paragraphs AND tables; location.paragraph counts paragraphs "
            "only).",
        ],
    },
    "bulk-edit": {
        "summary": (
            "Many edits across one document in as few calls as possible."
        ),
        "steps": [
            {"tool": "get_document_view",
             "why": "one call returns the text plus stable anchor ids for "
                    "every block (scope to one section to save tokens)"},
            {"tool": "apply_edits",
             "why": "apply the whole edit list in one transaction (one "
                    "lock, one backup, one save), addressed by the view's "
                    "anchors; a stale anchor refuses the whole batch"},
            {"tool": "search_and_replace",
             "why": "pattern edits across the document; preview=true "
                    "dry-runs the batch and yields the count for "
                    "max_replacements"},
            {"tool": "set_paragraph_text",
             "why": "surgical fallback for single-paragraph rewrites when a "
                    "find string would be unwieldy"},
            {"tool": "validate",
             "why": "structural check after the pass"},
        ],
        "notes": [
            "Heuristic: three or more text edits in one section, use view "
            "+ apply_edits; otherwise the fine-grained tools.",
            "Every mutation still auto-backs up, so a bad batch rolls back "
            "via manage_backups restore.",
        ],
    },
}


def _pack_of(tool_name: str) -> str | None:
    """The owning pack for a non-lite tool, from the live registry
    (discoverability rule 4: recipes name the packs their steps need)."""
    from .. import packs

    pack = packs.pack_of(tool_name)
    return None if pack in (None, "lite") else pack


def _packs_required(wf: dict) -> list[str]:
    return sorted({
        p for step in wf["steps"]
        if (p := _pack_of(step["tool"])) is not None
    })


def get_workflows(task: str | None = None) -> dict:
    """No task: list available tasks with summaries. With a task: the
    recommended tool sequence, one why-line per step, plus notes. Steps
    whose tool lives in an optional pack carry that pack's name, and
    packs_required lists every pack the workflow needs enabled."""
    if task is None:
        return {
            "tasks": [
                {
                    "task": name,
                    "summary": wf["summary"],
                    "packs_required": _packs_required(wf),
                }
                for name, wf in WORKFLOWS.items()
            ],
            "note": "call again with task='<name>' for the step-by-step sequence",
        }
    wf = WORKFLOWS.get(task)
    if wf is None:
        raise WordMcpError(
            f"unknown task {task!r}; available tasks: {sorted(WORKFLOWS)}"
        )
    steps = []
    for step in wf["steps"]:
        out = dict(step)
        pack = _pack_of(step["tool"])
        if pack is not None:
            out["pack"] = pack
        steps.append(out)
    required = _packs_required(wf)
    result = {"task": task, **wf, "steps": steps}
    result["packs_required"] = required
    if required:
        result["note"] = (
            f"steps tagged with a pack need enable_tools(packs="
            f"{required}) first; the rest run from the lite core"
        )
    return result
