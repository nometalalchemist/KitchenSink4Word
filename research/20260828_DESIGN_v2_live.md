# v2.0 Design — Live COM Editing (L0 output)

Basis: `20260828_L0_live_com_research.md` (8 topics, recommendation blocks) plus
local empirical probes run 2026-08-28 01:15 KST on this machine's Word 365.

## Empirical probe results (this Word build)

1. **UndoRecord empty-stack + wdReplaceAll: NO crash.** Both no-match and
   matching replace-all inside a custom record on a force-emptied undo stack
   survived. The reported Word bug does not reproduce here. Keep the
   try/except degradation anyway (`undo_grouped: false` fallback).
2. **Out-of-process undo grouping: VERIFIED.** Three Range edits in one custom
   record → a single `doc.Undo(1)` reverted all three.
3. **ScreenUpdating does NOT self-heal**: after a hard client kill
   (`os._exit`) with `ScreenUpdating = False`, the property stayed False in
   the surviving Word; a fresh attach repaired it. → watchdog + repair tool
   are mandatory; default is to NOT toggle ScreenUpdating except on >20-mutation
   batches.
4. **GetActiveObject attaches immediately** to a DispatchEx-spawned visible
   instance — no ROT registration lag observed in this pattern.
5. **IME composition probe: DEFERRED** (needs the user actively composing
   Korean). Bounded busy-retry (3×250ms) covers the expected RETRYLATER case;
   verify with the user in the morning.

## Architecture

**New module `src/word_mcp/com/live.py`** (sibling of bridge.py; bridge is the
batch layer with its own invisible instances — live.py NEVER creates instances,
never quits/saves/closes without an explicit flag).

Core pieces (all per the research doc's DESIGN SUMMARY):
- `live_session(path, tool_name)` contextmanager: CoInitialize → GetActiveObject
  (fail + WINWORD present → "unattachable" hint; else WordNotRunning) → doc
  resolve by FullName → ROT-scan fallback (multi-instance) → ProtectedViewWindows
  check → `probe_ready` (app.Name/doc.Name, 3×250ms retry on
  RPC_E_CALL_REJECTED/-2147418111, RPC_E_SERVERCALL_RETRYLATER/-2147417846) →
  StateGuard + `undo_group` → yield (app, doc, guard) → finally: end record
  (IsRecordingCustomRecord-guarded), guard.restore(), null refs, CoUninitialize.
- `StateGuard`: LIFO snapshot-on-mutate. Always-restore: TrackRevisions,
  UserName/UserInitials, app.Options.* touched, ScreenUpdating. Never-write:
  Visible, Saved, AutoSaveOn, DisplayAlerts (live instance), window/view state,
  clipboard.
- Error types (core/errors.py): WordNotRunning, DocumentNotOpenInWord,
  ProtectedViewRefused, WordBusy, WordBlocked, WordDisconnected
  (RPC_E_DISCONNECTED/-2147417848, CO_E_OBJNOTCONNECTED/-2147220995; message
  warns "edit may be partially applied — Ctrl+Z undoes the partial step").
- `probe_with_timeout` (helper-thread fresh attach, 5s) → 'blocked' detection;
  surfaced in com_word_status as `interactive_state`.
- `word_live_repair()` tool: fresh attach; ScreenUpdating=True; DisplayAlerts
  =wdAlertsAll; end orphaned custom records.
- Selection NEVER read or written (exception: live_insert_at_cursor READS
  Selection.Range.Start once; never writes). No .Select(), no ActiveWindow
  state, no clipboard tools (FormattedText assignment instead).
- Hardening from word-mcp-live (MIT — patterns re-implemented): 30KB text
  chunking, \x07 rejection, replace-loop guards (zero-length match,
  TrackRevisions off during replace_all with restore), inflated-range Find
  fallback, name[:64] truncation on undo records.

**Standard live result fields:** `live: true`, `undo_grouped`, `document_dirty`
(not doc.Saved), `autosave_on` (try/except-read), `state_restore_failed`,
`had_unsaved_user_changes` (pre-edit). `doc.Save()` only under explicit
`save=True`.

## Routing (L2)

File-based behavior untouched and default. Routed tools gain `live: str = "auto"`:
- "auto": try file path; on DocumentLocked → live implementation.
- "force"/True: straight to live (doc must be open in Word).
- "off"/False: current v1 refusal behavior.
Routed set (L2): search_and_replace, insert_paragraphs, delete_paragraphs,
set_cells, format_text, get_text, find_text, get_outline, get_document_info.
Live implementations in `com/live_ops.py`, addressing by paragraph index /
Find, mirroring file-based parameter shapes and result schemas (+ live fields).

## Live-specific tools (L3)

- `live_insert_at_cursor(text, ...)` — reads Selection.Range.Start once,
  inserts via Range (never touches Selection object afterward).
- `live_scroll_to(query|paragraph)` — the ONE sanctioned scroll:
  ActiveWindow.ScrollIntoView(rng) WITHOUT selecting; explicit user-facing
  "show me" tool.
- `live_set_track_changes(on/off)` — deliberate persistent toggle (no restore;
  this tool's purpose IS the state change).
- `word_live_repair()` — recovery.
- `com_word_status` extended: interactive_state ready|busy|blocked|not_running,
  per-doc dirty/autosave flags.

## Testing

- Live tests need a running Word → pytest marker `live`, auto-skip when Word
  absent/CI (CI runners have no Office; the 344 file-based tests keep CI green).
- Local harness spawns a VISIBLE DispatchEx instance with probe docs
  (only-test-docs guard before any op; user's real files unreachable).
- L4 bug-hunt: live × file-based × protection × fields, Selection preservation,
  undo grouping, crash-state restoration, mixed live/file sequences.
- L5 insane stress through raw MCP stdio transport with Word open.

Version target: 2.0.0. Ship gate: L4 + L5 green, all findings fixed.
