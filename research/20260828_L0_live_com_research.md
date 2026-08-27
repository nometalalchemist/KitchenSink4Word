# KitchenSink4Word v2.0 — Live COM Editing Research (L0)

Date: 2026-08-28 (KST). Scope: the eight live-COM topics from the v2 kickoff brief — editing documents OPEN in the user's interactive Word window. OOXML structures and batch-COM signatures are out of scope (covered in `20260827_v12_research.md`; its Topic 14 has the verified Compare/Merge/SaveAs2/proofing/InsertFile signatures).

Verification tags used below:
- **[MS-DOCS]** — stated on Microsoft Learn VBA reference (authoritative for the object model).
- **[REF-IMPL]** — observed in the MIT-licensed `ykarapazar/word-mcp-live` source (empirical evidence from a shipping implementation, not documentation).
- **[CONVERGENT]** — multiple independent non-Microsoft sources agree; no single authoritative statement.
- **[UNVERIFIED — LOCAL TEST]** — needs an empirical test on this machine; the exact test is given.

Existing foundation this builds on: `src/word_mcp/com/bridge.py` — per-call `pythoncom.CoInitialize()`/`CoUninitialize()`, `DispatchEx` for invisible instances, `GetActiveObject("Word.Application")` in `_find_open_document`, Quit-in-finally. The live layer keeps that per-call discipline and never calls Quit.

---

# TOPIC 1 — Range-based editing without touching Selection

## 1.1 The core premise holds

Every operation in the planned live tool surface is expressible through `Document.Range` / `Range` objects with no `Selection` involvement:

| Need | Range-based API |
| --- | --- |
| Address by character offset | `doc.Range(start, end)` |
| Address by paragraph | `doc.Paragraphs(i).Range` (1-based) |
| Address by bookmark | `doc.Bookmarks(name).Range` |
| Insert text | `rng.InsertBefore` / `rng.InsertAfter` / `rng.Text = ...` |
| Rich copy between locations | `rng.FormattedText = other_rng.FormattedText` |
| Find / replace | `rng.Find.Execute(...)` (full parity with Selection.Find, incl. `Replace:=wdReplaceAll`) |
| Formatting | `rng.Font`, `rng.ParagraphFormat`, `rng.Style` |
| Clipboard | `rng.Paste`, `rng.PasteAndFormat`, `rng.PasteSpecial` [MS-DOCS: Range.Paste exists; the range expands to cover pasted content, unlike Selection.Paste which lands after it] |
| Breaks / files | `rng.InsertBreak`, `rng.InsertFile` |
| Fields | `rng.Fields.Add`, `rng.Fields.Update` |
| Comments | `doc.Comments.Add(Range=rng, Text=...)` |
| Notes | `rng.Footnotes.Add` / `rng.Endnotes.Add` |
| Tables | `doc.Tables(i)`, `Tables.Add(Range=rng, ...)`, `table.Rows.Add`, `table.Columns.Add`, `cell.Range` |
| Headers/footers | `section.Headers(wdHeaderFooterPrimary).Range` (never `View.SeekView`) |
| Navigation targets | `rng.GoTo(...)` (returns a Range, does NOT scroll or select) |
| Position info | `rng.Information(wdActiveEndPageNumber etc.)` |

Two facts make this safe for the user's cursor:

1. **Word tracks the user's selection as an internal range and adjusts it automatically when text before/around it changes** — exactly as it adjusts bookmarks and other ranges. An insertion at offset 100 while the user's cursor sits at offset 500 leaves the cursor on the same *text*, now at offset 500+n. No restoration needed. [CONVERGENT — this is the entire premise of every co-authoring and add-in edit; also empirically how `word-mcp-live` operates 44 live tools with zero Selection code, [REF-IMPL]]
2. **Range operations do not scroll the window** and do not move focus, as long as you never call `rng.Select()` or `ActiveWindow.ScrollIntoView(rng)`.

## 1.2 Operations that genuinely require Selection

Short list, and none of them are needed for the planned v2 surface:

- **`Selection.CopyFormat` / `Selection.PasteFormat`** (format painter). No Range equivalent. Workaround that avoids Selection entirely: read the source range's `Font`/`ParagraphFormat` properties and assign them to the target range property-by-property, or use `rng.Font = other.Font.Duplicate` / `rng.ParagraphFormat = other.ParagraphFormat.Duplicate` (the `.Duplicate` assignment pattern is the documented way to transfer format objects).
- **Typing emulation** — `Selection.TypeText`, `TypeParagraph`, `TypeBackspace`. Only needed to simulate keystrokes; `InsertBefore/After` covers real use.
- **Built-in dialogs** — `app.Dialogs(wdDialog...).Show/Execute` mostly operate on the Selection. Do not ship dialog-based tools in a live layer.
- **Caret/extend navigation** — `Selection.EndKey`, `HomeKey`, `Extend`, `Shrink`, `MoveLeft`... irrelevant to programmatic editing.
- **A few paste variants** — `Selection.PasteAppendTable` has no Range twin (`Range.PasteExcelTable` and `Range.PasteAndFormat` DO exist). Avoid the tool; if ever needed, use the save/restore pattern below.
- **View state** — anything that changes what the user is looking at (`ActiveWindow.View.Type`, `SeekView`, zoom, `PrintPreview`). A live layer must simply never touch these.

Conclusion: **design rule — the live layer never reads or writes `app.Selection` at all**, except inside the one sanctioned wrapper below.

## 1.3 Save/restore pattern for the unavoidable case

If a future tool must use Selection, the naive snapshot (`s, e = sel.Start, sel.End` ... `doc.Range(s, e).Select()`) breaks whenever the edit inserts/deletes text before the saved offsets. The robust pattern anchors with a **temporary bookmark**, because Word adjusts bookmarks through edits:

```python
def selection_guard(app, doc):
    """Use ONLY when a Selection-dependent op is unavoidable."""
    sel = app.ActiveWindow.Selection
    BM = "_wmcp_selguard"
    doc.Bookmarks.Add(BM, sel.Range)          # zero-length bookmarks allowed
    try:
        yield sel
    finally:
        try:
            if doc.Bookmarks.Exists(BM):
                doc.Bookmarks(BM).Range.Select()
                doc.Bookmarks(BM).Delete()
            else:
                # edge case: the edit deleted the bookmarked span entirely.
                # Clamp to a nearby valid position instead of failing.
                pos = min(sel_start_saved, doc.Content.End - 1)
                doc.Range(pos, pos).Select()
        except Exception:
            pass  # never let cursor restore mask the real result
```

Edge cases the bookmark handles / doesn't:
- Edit before the selection → bookmark shifts, restore is exact. Handled.
- Edit *inside* the selection → bookmark stretches with it (Word's bookmark semantics). Handled, usually desirable.
- Edit deletes the whole bookmarked span → **bookmark is deleted**; `Bookmarks.Exists` returns False → clamp fallback.
- Selection in another story (footnote pane, header) → bookmarks work in all stories; `Range.Select()` re-selects in that story, but if the op switched the active pane the user's pane focus may still change. This is why Selection ops stay banned by default.

**RECOMMENDATION (Topic 1).** All live tools address content exclusively via `doc.Range(start, end)`, `doc.Paragraphs(i).Range`, bookmarks, and `rng.Find.Execute`. Adopt word-mcp-live's inflated-range defense [REF-IMPL]: when a paragraph's `Range.End` is inflated by embedded comments/fields, fall back to `Find` + `Expand` to locate the true paragraph mark. Never call `.Select()`, never read `app.Selection`, never change `ActiveWindow` state. Ship `selection_guard` in the codebase but leave it unused; document that any tool wanting it needs an explicit design review.

---

# TOPIC 2 — Application.UndoRecord: one tool call = one Ctrl+Z step

## 2.1 API surface [MS-DOCS]

Source: "Working with the UndoRecord Object" (Word VBA concepts, ms.date 2019-06-08) and `word.undorecord.startcustomrecord`.

```python
undo = app.UndoRecord                       # property of Application
undo.StartCustomRecord("word-mcp: replace text")   # Name optional, shows in the Undo dropdown
# ... any number of object-model edits ...
undo.EndCustomRecord()
```

- `StartCustomRecord([Name])` — Name optional; if omitted, Word names the record after the first command executed. No documented length limit; **word-mcp-live truncates to 64 chars** (`rec.StartCustomRecord(name[:64])`) [REF-IMPL] — adopt the truncation defensively.
- `EndCustomRecord()` — no args.
- `IsRecordingCustomRecord` (Boolean, read-only) — use this to guard the finally-block End call.
- `CustomRecordLevel` (Long) — number of active nested Start calls; nesting is legal, only the outermost creates the stack entry.
- `CustomRecordName` (String, read-only).

## 2.2 Constraints — what breaks a custom record [MS-DOCS]

Verbatim-grade facts from the concepts page:

1. **Switching documents inside a record terminates the record.** "Word terminates the custom undo record when the code begins to write to the second document." → a live tool that edits two documents cannot get one combined undo step; scope every record to exactly one document.
2. **`Document.Undo` called inside a record** undoes the record's prior actions and splits behavior oddly — "Calling the Undo method in the wrong order within a custom undo record can have undesired effects." Never call `doc.Undo` inside a record.
3. **Word auto-ends unclosed records**: "Close each custom undo record with a call to EndCustomRecord. Word will attempt to determine where to end the record, but it may not be at the desired point of code execution." This is the crash-safety answer: **if the client process dies mid-record, Word ends the record itself at the next boundary** — the user's Word does not get stuck in a recording state, but the grouping boundary is then arbitrary. So: End in a finally, always; a crash degrades gracefully to "wrong grouping," not "broken Word."
4. Debugger breakpoints (VBE) auto-end records — irrelevant out-of-process, but confirms 3's mechanism.

Undo-stack clearing: the stack is cleared explicitly by `doc.UndoClear()` (never call it) and implicitly by closing the document. `Document.Save` does NOT clear the undo stack. No documented list of "operations that clear the stack" exists beyond these; treat any tool that observes an emptied stack as a bug to investigate. [MS-DOCS for UndoClear; rest CONVERGENT]

## 2.3 Known bug — crash with empty undo stack

Microsoft Q&A 428399 ("Word VBA. UndoRecord.StartCustomRecord crush Word"): **`StartCustomRecord` crashes Word when the undo list is empty**, reported specifically in combination with `Find.Execute Replace:=wdReplaceAll` inside the record. Present since Word 2010, still reported 2021, no Microsoft fix or workaround in the thread. [CONVERGENT — single Q&A + "mentioned on the internet"; severity is high (host crash) so treat as real until disproven]

**[UNVERIFIED — LOCAL TEST]** Run on this machine's Word build before shipping:
```python
# test_undorecord_empty_stack.py — run with a THROWAWAY doc open in visible Word
import pythoncom, win32com.client
pythoncom.CoInitialize()
app = win32com.client.GetActiveObject("Word.Application")
doc = app.ActiveDocument            # fresh doc: undo stack is empty
app.UndoRecord.StartCustomRecord("crash probe")
rng = doc.Content
f = rng.Find
f.Text = "zzzz_not_present"; f.Replacement.Text = "x"
f.Execute(Replace=2)                # wdReplaceAll = 2
app.UndoRecord.EndCustomRecord()
print("survived; IsRecording =", app.UndoRecord.IsRecordingCustomRecord)
```
If Word crashes: mitigation is to seed the stack before starting the record when `doc` is pristine — an easy detectable proxy is "document has no undoable action yet"; seeding with a no-op undoable edit (insert+delete a space at doc start inside its own mini-record) is ugly; the cleaner mitigation is to skip the custom record for `replace_all` on fresh documents and accept multiple undo entries there, reporting `"undo_grouped": false`.

## 2.4 Out-of-process behavior

The object model does not restrict `UndoRecord` to in-process callers, and **word-mcp-live drives it from an out-of-process Python client as its standard per-tool pattern, with per-call undo grouping reported working** ("Per-action Ctrl+Z on Windows", `with undo_record(app, "MCP: Insert Text")`) [REF-IMPL]. It also catches the exception `StartCustomRecord` raises on Word 2007 (API added in 2010) and degrades to ungrouped edits. Treat out-of-process operation as working; confirm once with the local test above (which is itself out-of-process).

**RECOMMENDATION (Topic 2).**
```python
import contextlib

@contextlib.contextmanager
def undo_group(app, name: str):
    undo = app.UndoRecord
    started = False
    try:
        undo.StartCustomRecord(name[:64])
        started = True
    except Exception:
        pass                          # old Word / API refusal: degrade, don't fail the edit
    try:
        yield started                 # tool reports "undo_grouped": started
    finally:
        if started:
            with contextlib.suppress(Exception):
                if undo.IsRecordingCustomRecord:
                    undo.EndCustomRecord()
```
Rules: one record per tool call, opened AFTER the document is resolved and any state snapshot is taken, closed in finally before state restore; never span two documents; never call `doc.Undo`/`doc.UndoClear` inside; run the empty-stack crash test locally before v2 ships and gate `replace_all` accordingly.

---

# TOPIC 3 — ScreenUpdating: performance toggle with guaranteed restoration

## 3.1 What the docs guarantee (and don't) [MS-DOCS]

`Application.ScreenUpdating` (Word VBA reference): "controls most display changes on the monitor while a procedure is running... **You must set the ScreenUpdating property to True when the procedure finishes or when it stops after an error.**" Companion method `Application.ScreenRefresh` forces one repaint while updating is off.

Critically, **there is NO documented auto-restore** — not at macro end, not when a COM client releases its references or disconnects. The docs put the burden entirely on the caller. Community reporting cuts both ways: several Microsoft Q&A threads report modern Word 365 builds *ignoring or aggressively re-enabling* `ScreenUpdating = False` (a behavior change complained about post-2021 updates), which would make a stuck-frozen Word unlikely on current builds — but that is the opposite failure (the optimization silently stops working), and neither behavior is documented. [CONVERGENT, contradictory reports]

**[UNVERIFIED — LOCAL TEST]** Definitive answer for this machine's build, 2 minutes:
```python
# test_screenupdating_disconnect.py — visible Word with a throwaway doc open
import pythoncom, win32com.client, os
pythoncom.CoInitialize()
app = win32com.client.GetActiveObject("Word.Application")
app.ScreenUpdating = False
app.ActiveDocument.Content.InsertAfter("frozen? ")
os._exit(1)   # simulate client crash: no restore, no CoUninitialize, no cleanup
```
Then observe: (a) does the Word window still repaint when you type into it? (b) if frozen, does typing/clicking un-freeze it? (c) run a second script that re-attaches and reads `app.ScreenUpdating` — does it report False? Record the result in BUILD_LOG; it decides how paranoid the watchdog must be.

## 3.2 The restoration discipline

Regardless of the local test result, the pattern is:

```python
prev = None
try:
    prev = app.ScreenUpdating         # snapshot (user may already have it off via another add-in)
    app.ScreenUpdating = False
    ...edits...
finally:
    try:
        app.ScreenUpdating = True if prev is None else prev
        app.ScreenRefresh()           # force immediate repaint
    except Exception:
        _schedule_screen_repair()     # last resort, below
```

**Last-resort watchdog.** The property lives in the Word process, so a *fresh attachment can always repair it* — you do not need the original COM pointer. Two layers:

1. **In-process retry**: if the finally-restore raises (RPC error because Word was busy at that instant), spawn a daemon thread that sleeps 1s, does its own `CoInitialize` + `GetActiveObject("Word.Application")`, sets `ScreenUpdating = True`, retries ×3. Covers transient rejection.
2. **Repair tool**: ship `word_live_repair()` — attaches fresh and sets `ScreenUpdating = True`, `DisplayAlerts = wdAlertsAll (-1)`, and ends any orphaned custom undo record (`if app.UndoRecord.IsRecordingCustomRecord: EndCustomRecord()`). The user (or the model, on the next tool call after an error) can always invoke it. This covers the hard-crash case where no finally ran at all.

Note `zombie_check()` already exists in bridge.py; `word_live_repair` is its live-instance sibling.

**RECOMMENDATION (Topic 3).** Do NOT toggle ScreenUpdating by default. word-mcp-live ships 44 live tools without it and single-Range edits repaint imperceptibly [REF-IMPL]; the freeze risk buys nothing on small edits. Toggle it ONLY in tools that perform many sequential COM edits (batch replace across stories, N-row table fills — threshold: more than ~20 discrete COM mutations), always via the snapshot/finally pattern above, plus the daemon-thread retry and the `word_live_repair` tool. Run the disconnect test locally and record whether this Word build self-heals.

---

# TOPIC 4 — Interactive-instance state: what the live layer may mutate, and the snapshot/restore discipline

## 4.1 Complete inventory of state a live-editing layer might touch

**Application-level (shared with the user's session — highest blast radius):**

| Property | Why a tool touches it | Default / restore target |
| --- | --- | --- |
| `app.ScreenUpdating` | batch performance (Topic 3) | snapshot; effectively True |
| `app.DisplayAlerts` | suppress prompts during an edit | `wdAlertsAll = -1` is Word's default; snapshot — but see 4.3: prefer NOT suppressing on the user's instance |
| `app.UserName` / `app.UserInitials` | tracked-change/comment attribution ("Claude" vs the user) | snapshot & restore ALWAYS if changed — word-mcp-live does this per-tool [REF-IMPL] |
| `app.Options.CheckGrammarWithSpelling` | proofing pass | snapshot |
| `app.Options.ReplaceSelection`, other `app.Options.*` | only if a tool sets them | snapshot each one it sets |
| `app.Visible` | **NEVER touch on the user's instance** | n/a |
| `app.WindowState`, `ActiveWindow.*`, `View.*`, zoom | **NEVER touch** | n/a |

**Document-level (scoped to the target doc):**

| Property | Why | Restore |
| --- | --- | --- |
| `doc.TrackRevisions` | force tracked/untracked edits per tool flag; word-mcp-live also disables it during `replace_all` to stop Find re-matching visible deletions [REF-IMPL] | snapshot & restore ALWAYS |
| `doc.SpellingChecked` / `doc.GrammarChecked` | reset to force re-proof (bridge.py already does on invisible instances) | snapshot; note resetting dirties proofing state |
| `doc.ShowSpellingErrors` / `doc.ShowGrammaticalErrors` | hide squiggle churn during proofing pass | snapshot |
| `doc.Saved` | READ-ONLY for us (dirty detection). **Never write** — setting `Saved = True` suppresses the close-prompt and can silently lose user work | n/a |
| `doc.AutoSaveOn` | READ-ONLY for us (Topic 5 reporting). Never write silently | n/a |

**Not state, but session-visible side effects to avoid:** adding to `RecentFiles` (only applies to Open, which the live layer never does), clipboard contents (avoid `Copy`/`Paste`-based tools; use `FormattedText` assignment instead — the user's clipboard is user state), and the undo stack itself (Topic 2 handles grouping; never `UndoClear`).

## 4.2 The snapshot/restore discipline

One guard object, declared per tool, restore in reverse order, tolerate partial failure and *report* it rather than mask it:

```python
class StateGuard:
    """Snapshot exactly the properties a tool declares it will touch."""
    def __init__(self):
        self._stack = []      # (obj, attr, saved_value)

    def set(self, obj, attr, value):
        self._stack.append((obj, attr, getattr(obj, attr)))
        setattr(obj, attr, value)

    def restore(self) -> list[str]:
        failed = []
        for obj, attr, saved in reversed(self._stack):
            try:
                setattr(obj, attr, saved)
            except Exception:
                failed.append(attr)
        self._stack.clear()
        return failed          # goes into the tool result as state_restore_failed
```

Usage inside the session wrapper: `guard.set(doc, "TrackRevisions", False)` — the guard records the prior value at mutation time, so nothing is snapshotted that isn't touched, and restore order is automatically LIFO. Restore runs in the session finally AFTER `EndCustomRecord` (so the restores are not recorded in the user's undo step — property changes like TrackRevisions are not undoable edits anyway, but ordering keeps the record clean) and BEFORE `CoUninitialize`.

## 4.3 DisplayAlerts on the user's instance — a deliberate deviation from bridge.py

bridge.py sets `DisplayAlerts = 0` on its own invisible instances, correctly. On the USER's instance the calculus flips: suppressing alerts globally while the user is simultaneously working suppresses THEIR alerts too (e.g., their own save-conflict prompt firing mid-edit). Default: leave `DisplayAlerts` alone; live tools should instead be written so they never trigger prompts (no Open, no Close, no SaveAs, no format conversions on the live path). If a specific tool provably needs it, it declares it through the guard.

**RECOMMENDATION (Topic 4).** Ship `StateGuard`; every live tool mutates app/doc state only through `guard.set()`. Always-restored set: `TrackRevisions`, `UserName`/`UserInitials` (when attribution is changed), anything in `app.Options`. Never-write set: `Visible`, `Saved`, `AutoSaveOn`, all window/view state, clipboard. `DisplayAlerts` stays untouched by default on the live instance. `state_restore_failed` is a first-class field in every live tool's result schema.

---

# TOPIC 5 — Dirty-document semantics: composing with the user's unsaved edits

## 5.1 The intended model is sound

COM edits against an open document go through exactly the same editing pipeline as keystrokes: they land in the in-memory document, join the same undo stack, set `doc.Saved = False`, and compose with whatever unsaved edits the user has already made. There is no second buffer and no conflict between "the user's unsaved changes" and "the tool's changes" — they are the same change stream. The user saves when they choose; `doc.Save()` is called only under an explicit `save=True` tool flag (mirroring the existing `com_save_open_document`). This is precisely how VSTO add-ins and word-mcp-live operate. [CONVERGENT + REF-IMPL]

Corollary: the file on disk is STALE the moment either the user or a tool edits. The existing rule stands — file-based (OOXML) tools must keep refusing open documents; the live layer is the only writer while Word holds the doc, and reading current content must happen via COM (`doc.Content.Text` etc.) or after an explicit save.

## 5.2 Gotchas

1. **AutoSave documents (OneDrive/SharePoint).** `Document.AutoSaveOn` (read/write Boolean) [MS-DOCS: `word.document.autosaveon`] — defaults **True for cloud-hosted docs**, False locally. When True, Word saves continuously: the "user saves when ready" model does not hold — tool edits hit the server copy within seconds, and "undo" only rolls back the local doc (which then autosaves the rollback — consistent, but the user should know edits were persisted immediately). **Pattern: every live tool reads `doc.AutoSaveOn` (wrapped in try/except — property may raise on older builds/formats) and reports `"autosave_on": true/false` in its result.** Never toggle it silently; if the user wants a "sandbox" edit session on a cloud doc, that is a user decision surfaced through a dedicated flag that toggles-and-restores via StateGuard, with the caveat that Microsoft warns AutoSave can also be turned off/on by co-authoring transitions.
2. **Co-authoring.** On a co-authored doc, COM edits merge through the same co-authoring pipeline as typing. Conflicts are possible if the tool edits a paragraph another author holds. Detection surface: `doc.CoAuthoring.Authors.Count > 1` indicates active co-authors. [UNVERIFIED — LOCAL TEST if co-authoring support is ever prioritized; low priority for a single-user machine. Report the count when > 1, refuse nothing.]
3. **AutoRecover background saves** happen on a timer in the user's instance and briefly make Word busy; a COM call landing in that window gets `RPC_E_SERVERCALL_RETRYLATER` — handled by the Topic 7 bounded retry, not a special case.
4. **`doc.Saved` manipulation** — repeated for emphasis: never set it. Reading it is useful (`"had_unsaved_user_changes": not doc.Saved` before editing is honest telemetry for the tool result).
5. **Zero-net-change tools must not dirty the doc.** Read-only live tools (outline read, text read) must avoid calls with edit side effects (`Fields.Update`, `Repaginate`, resetting `SpellingChecked`). If a "read" tool must run one of these, it is not a read tool — say so in its description.

**RECOMMENDATION (Topic 5).** Keep the model: in-memory edits, user saves, `save=True` flag calls `doc.Save()` explicitly. Add to every live tool result: `"document_dirty": not doc.Saved` (post-edit), `"autosave_on": <bool|None>`. Never write `Saved` or `AutoSaveOn`. Keep the file-based lock refusal exactly as is; add the inverse guard to live tools (refuse if the doc is NOT open in Word — already the `_find_open_document` behavior).

---

# TOPIC 6 — COM marshaling & threading from a long-lived stdio server

## 6.1 The threading reality of this server

pywin32's `CoInitialize()` enters a **single-threaded apartment (STA)** — correct for Office automation. STA rules give every COM interface pointer **thread affinity**: a pointer obtained on thread A may not be used from thread B without explicit marshaling (`CoMarshalInterThreadInterfaceInStream` / Global Interface Table), which pywin32 does not do automatically.

The MCP server is a long-lived asyncio (FastMCP/anyio) process, and **sync tool functions execute on worker threads from a pool — the same OS thread is NOT guaranteed across tool calls.** Therefore:

- **Caching `app`/`doc` COM objects across tool calls is unsafe twice over:** (a) the next call may run on a different pool thread → illegal cross-apartment use (symptoms range from working-by-luck to `CoInitialize has not been called`, marshaling errors, or corruption); (b) Word can quit between calls → the cached proxy raises `RPC_E_DISCONNECTED`.
- **Per-call attach is the right pattern, and it is cheap.** `GetActiveObject` is a Running Object Table lookup — on the order of a millisecond, noise against any actual edit. bridge.py's per-call `CoInitialize`/`CoUninitialize` discipline is exactly right and extends unchanged to the live layer.

One correctness refinement over the current bridge.py: **release COM references before `CoUninitialize`.** Calling `CoUninitialize` while proxy objects are still alive risks crashes at teardown. Null the Python variables (and drop any locals holding Ranges/Documents) first:

```python
pythoncom.CoInitialize()
app = doc = None
try:
    app = win32com.client.GetActiveObject("Word.Application")
    doc = _match_document(app, path)
    ...work...
finally:
    doc = None      # release proxies BEFORE CoUninitialize
    app = None
    pythoncom.CoUninitialize()
```
(CPython refcounting frees the PyIDispatch immediately on rebind; no gc pass needed unless ranges are captured in closures — don't do that.)

`CoInitialize` on an already-initialized STA thread returns S_FALSE harmlessly; the pairing must stay balanced per call, which the pattern above does.

## 6.2 GetActiveObject failure modes and the ROT

- **`MK_E_UNAVAILABLE` (0x800401E3, com_error -2147221021, "Operation unavailable")** — no Word in the Running Object Table. Usual meaning: Word is not running. BUT two real gotchas [CONVERGENT; the old MS KB on this is now 410-gone but its content is widely mirrored]:
  1. **ROT registration lag**: an Office app may not register in the ROT until it loses focus once after startup. A Word launched seconds ago and never de-focused can be invisible to `GetActiveObject`.
  2. **Integrity-level mismatch**: an elevated client cannot see a non-elevated Word's ROT entry and vice versa. The MCP server runs non-elevated like Word, so this only bites if the user elevates one side.
  Detection: if `GetActiveObject` fails but `zombie_check()` sees WINWORD.EXE, report "Word is running but not attachable (just launched and not yet registered, or elevation mismatch) — click into another window once, then retry."
- **Multiple Word instances**: `GetActiveObject("Word.Application")` returns whichever instance registered; the target doc may be in another. word-mcp-live's answer [REF-IMPL] is a **ROT scan fallback**: enumerate the ROT for file monikers ending in the target path and bind straight to the Document:

```python
def _find_doc_via_rot(path_lower: str):
    rot = pythoncom.GetRunningObjectTable()
    for moniker in rot.EnumRunning():
        ctx = pythoncom.CreateBindCtx(0)
        try:
            name = moniker.GetDisplayName(ctx, None)
        except Exception:
            continue
        if name.lower() == path_lower:
            obj = rot.GetObject(moniker)
            doc = win32com.client.Dispatch(
                obj.QueryInterface(pythoncom.IID_IDispatch))
            return doc, doc.Application       # Document -> its own Application
    return None, None
```
  Open documents register their full path as a file moniker, so this finds the exact instance holding the doc. Use it as the fallback when the `GetActiveObject` instance doesn't contain the target path.
- **`RPC_E_DISCONNECTED` (0x80010108, -2147417848, "The object invoked has disconnected from its clients")** and **`CO_E_OBJNOTCONNECTED` (0x800401FD, -2147220995)** — the server died under a live proxy (Word quit or crashed mid-call). With per-call attach these can still occur *mid-tool* (user closes Word while an edit runs). Handling: catch at the session wrapper, classify as `WordDisconnected`, attempt ONE silent re-attach + document re-resolve; if the doc is gone, surface "Word or the document was closed while editing — the edit may be partially applied" (honest, because a multi-step edit may have half-landed; the custom undo record, if Word survived, lets the user Ctrl+Z the partial edit).

**RECOMMENDATION (Topic 6).** Per-call lifecycle, no caching: `CoInitialize` → `GetActiveObject` → path-match in `app.Documents` → ROT-scan fallback → work → null refs → `CoUninitialize`. This is bridge.py's existing shape; add (a) ref-release before CoUninitialize, (b) the ROT-scan fallback for multi-instance, (c) the WINWORD-but-unattachable diagnostic, (d) `RPC_E_DISCONNECTED`/`CO_E_OBJNOTCONNECTED` classified as `WordDisconnected` with one re-attach retry. Do not introduce marshaling, GIT, or a dedicated COM thread — unnecessary complexity for per-call attach, and the per-call cost is ~1ms.

---

# TOPIC 7 — Modal / blocked states: detect and refuse, never hang

## 7.1 What actually happens when Word is modal

Word has **no `Application.Busy` or `.Ready` property** (Excel's `Application.Ready` has no Word counterpart — absent from the Word object model). Detection must be behavioral.

When the interactive Word has a modal dialog, Backstage (File menu / modern print preview), or is mid-command, its STA rejects incoming COM calls via the standard OLE message-filter mechanism: the server returns `SERVERCALL_REJECTED` or `SERVERCALL_RETRYLATER`, and the client-side call **fails fast** — it does not block — with one of:

- **`RPC_E_CALL_REJECTED`** = 0x80010001 = com_error hresult **-2147418111**, "Call was rejected by callee."
- **`RPC_E_SERVERCALL_RETRYLATER`** = 0x8001010A = **-2147417846**, "The message filter indicated that the application is busy."

[CONVERGENT — OLE architecture docs (IMessageFilter/HandleInComingCall), plus decades of Office-automation reports of exactly these two codes.] `GetActiveObject` itself usually still succeeds (ROT lookup doesn't enter Word's apartment the same way); the first *method/property* call is what gets rejected.

**Blocking forever is the rare case**: it happens only when Word is inside a long synchronous operation (huge repagination, printing, a synchronous add-in) — then the call queues and waits. There is no per-call timeout in COM.

## 7.2 IMessageFilter — considered and rejected

The classical fix is registering client-side `IMessageFilter::RetryRejectedCall` via `CoRegisterMessageFilter` to auto-retry rejected calls. **pywin32 does not expose `CoRegisterMessageFilter`** (confirmed on the python-win32 list; the third-party `imessagefilter` PyPI package exists precisely to fill this gap with a C extension). For this design it is also the *wrong* semantics: a message filter silently retries — possibly for the whole time a dialog sits open — while the requirement here is to **refuse cleanly and tell the model why**. A bounded try/except retry gives the same protection with no extension module and no process-global filter state.

## 7.3 The concrete detection-and-refusal pattern

```python
BUSY_HRESULTS = {-2147418111,   # RPC_E_CALL_REJECTED       (modal dialog / rejecting)
                 -2147417846}   # RPC_E_SERVERCALL_RETRYLATER (busy, retry later)
GONE_HRESULTS = {-2147417848,   # RPC_E_DISCONNECTED
                 -2147220995}   # CO_E_OBJNOTCONNECTED

def probe_ready(app, doc, retries=3, delay=0.25):
    """Cheap-call probe: classify the interactive instance's state."""
    import time
    for attempt in range(retries):
        try:
            _ = app.Name          # cheapest possible round-trip into Word's STA
            _ = doc.Name
            return                # ready
        except pywintypes.com_error as e:
            hr = e.hresult if e.hresult in BUSY_HRESULTS | GONE_HRESULTS \
                 else (e.args[2][5] if e.args[2] else e.hresult)
            if hr in GONE_HRESULTS:
                raise WordDisconnected("Word closed or crashed") from e
            if hr in BUSY_HRESULTS and attempt < retries - 1:
                time.sleep(delay); continue
            raise WordBusy(
                "Word is busy or has a dialog open (a dialog box, the File "
                "menu/Backstage, or a running command). Close it and retry."
            ) from e
```

Notes on the hresult plumbing: pywin32 surfaces the failing HRESULT either as `com_error.hresult` directly (call-level rejection) or nested in `excepinfo` for server-raised errors; check both, as above. Every live tool runs `probe_ready` right after attach and before starting the undo record — a rejected probe costs nothing and refuses before any mutation.

**Hard timeout for the block-forever case.** COM pointers are thread-affine, so you cannot offload an in-flight call to a watchdog. Instead, run the PROBE itself on a helper thread that does its own fresh attach:

```python
def probe_with_timeout(path, timeout=5.0) -> str:
    """'ready' | 'busy' | 'blocked' | 'not_running' — safe to call from any thread."""
    import threading
    result = {}
    def _worker():
        pythoncom.CoInitialize()
        try:
            app = win32com.client.GetActiveObject("Word.Application")
            _ = app.Name
            result["state"] = "ready"
        except pywintypes.com_error as e:
            result["state"] = "busy" if _is_busy(e) else "not_running"
        finally:
            pythoncom.CoUninitialize()
    t = threading.Thread(target=_worker, daemon=True)
    t.start(); t.join(timeout)
    return result.get("state", "blocked")   # no answer in time => Word not pumping
```
`'blocked'` → refuse with "Word is in the middle of a long operation (printing, repaginating a large document); wait for it to finish." The abandoned daemon thread's queued call completes or dies harmlessly with the process.

## 7.4 Specific states

| State | Detection | Action |
| --- | --- | --- |
| Modal dialog / Backstage / modern print preview | probe fails with REJECTED/RETRYLATER | refuse: "close the dialog" |
| Legacy print preview | `app.PrintPreview == True` (readable when not modal) | refuse politely; (do not set it False — user state) |
| **Protected View** | the doc is NOT in `app.Documents`; enumerate `app.ProtectedViewWindows(i).Document.FullName` | refuse: "document is in Protected View — click Enable Editing" |
| Word closing / closed mid-call | DISCONNECTED / OBJNOTCONNECTED | `WordDisconnected`, one re-attach attempt |
| Long synchronous op (not pumping) | `probe_with_timeout` returns 'blocked' | refuse: "wait for Word to finish" |
| IME composition in progress | **[UNVERIFIED — LOCAL TEST]**: with Korean IME mid-composition in the doc, run `probe_ready` from a second terminal; record whether calls are rejected, queued, or succeed (and whether an edit landing mid-composition disturbs the composition string) | expected: brief RETRYLATER at worst → covered by bounded retry; verify on this machine (Korean IME is in daily use here — this test matters) |

**RECOMMENDATION (Topic 7).** No message filter. Every live session = attach → `probe_ready` (3 tries × 250ms) → proceed or raise `WordBusy`/`WordDisconnected`/`ProtectedViewRefused` as typed, user-actionable errors. Expose `probe_with_timeout` inside `com_word_status` (report `interactive_state: ready|busy|blocked|not_running`) so the model can check before batching edits. Run the IME-composition test locally before v2 ships.

---

# TOPIC 8 — Reference implementation: ykarapazar/word-mcp-live

## 8.1 License — VERIFIED FIRST

GitHub license API for `ykarapazar/word-mcp-live`: **SPDX `MIT`, "MIT License"**, standard MIT text with copyright notices **"GongRzhe (2025)" and "Yüce Karapazar (2026)"** (it is a fork of GongRzhe/Office-Word-MCP-Server with a live-editing layer added). Permissive → source examined in full below.

## 8.2 Architecture (from source study: `word_document_server/core/word_com.py`, `tools/live_tools.py` ~109KB, `tools/live_read_tools.py`, `tools/live_layout_tools.py`)

**Attach:** `win32com.client.GetActiveObject("Word.Application")`, with a **ROT-scan fallback** (`_find_word_with_docs()`) that enumerates running-object-table monikers to find a Word instance actually holding documents when GetActiveObject fails or returns an empty instance. **No caching** — `get_word_app()` / `find_document(app, filename)` do fresh lookups every call. **No explicit CoInitialize** — relies on import-time initialization of the main thread; this is its most significant weakness for a threaded MCP server (works because their FastMCP configuration happens to keep calls on one thread; fragile — our per-call CoInitialize discipline is strictly better).

**Tool structure:** every live tool follows one skeleton: platform check → argument validation → `get_word_app()` + `find_document()` → `with undo_record(app, "MCP: <op>")` → save `doc.TrackRevisions` (and `app.UserName` where attribution changes), restore in finally → structured JSON result. 44 Windows live tools; macOS parity is partial (JXA stubs).

**Undo:** `undo_record(app, name)` contextmanager; `rec.StartCustomRecord(name[:64])`; catches the exception on Word 2007 and degrades to ungrouped. One record per tool call — their README's "per-action Ctrl+Z" claim is this mechanism working out-of-process.

**Selection:** never touched. All addressing is `doc.Range(start, end)`, `doc.Paragraphs(i).Range`, bookmarks, and `Find.Execute`. No selection save/restore exists because none is needed — direct confirmation of Topic 1's premise.

**Error handling & hardening worth adopting:**
- **Inflated-range fallback**: when comments/fields inflate a paragraph `Range.End`, they re-locate the true paragraph mark via Find + `Expand` instead of trusting offsets.
- **32K chunking**: text insertion chunked to ~30KB fragments to stay under Word COM's ~32K string limit per call.
- **Control-byte rejection**: refuses `\x07` (table cell separator) in plain-text input — inserting it corrupts table structure.
- **Replace-loop guards**: zero-length-match guard (infinite loop), 255-char replacement limit (Word's Find limit), and `TrackRevisions` forced off during `replace_all` (with restore) because visible tracked deletions re-match the Find pattern and loop forever.
- Escape-sequence preprocessing (`\\r\\n` → `\r` etc.) at the tool boundary.

**Weaknesses (do NOT copy):** no CoInitialize discipline; no ScreenUpdating management at all; no busy/modal detection (a dialog-open Word surfaces as a raw com_error to the model); the TrackRevisions save/restore is copy-pasted boilerplate across 15+ tools instead of a shared guard; no protected-view handling; loose result schemas.

**RECOMMENDATION (Topic 8).** Adopt: ROT-scan fallback, undo-record contextmanager (with our IsRecordingCustomRecord finally-guard added), inflated-range fallback, 30KB chunking, `\x07` rejection, replace-loop guards, and the general "one skeleton per tool" shape — but centralize it in ONE session contextmanager instead of per-tool boilerplate. Reject: their missing CoInitialize, missing busy detection, missing state-guard abstraction. Attribution: none required beyond MIT notice retention IF code is copied verbatim; the patterns above are being re-implemented, not copied, but if any function is lifted near-verbatim, carry the MIT notice per its terms.

---

# DESIGN SUMMARY — the recommended live-bridge architecture

**Module:** `src/word_mcp/com/live.py`, sibling to `bridge.py`, sharing its error types. bridge.py stays the batch layer (own invisible instances, Quit-in-finally); live.py NEVER creates instances and NEVER quits, saves, or closes anything without an explicit flag.

**Attach lifecycle (per tool call — no caching, ever):**
```
CoInitialize (STA)
 → GetActiveObject("Word.Application")
     fail + WINWORD.EXE present → "running but unattachable (focus/elevation)" refusal
     fail + no WINWORD.EXE      → WordNotRunning
 → resolve doc: FullName match in app.Documents
     miss → ROT scan by file moniker (finds the right instance in multi-instance setups)
     miss → check ProtectedViewWindows → ProtectedViewRefused, else DocumentNotOpenInWord
 → probe_ready (app.Name/doc.Name, 3×250ms on REJECTED/RETRYLATER) → WordBusy on refusal
 → StateGuard created; undo_group("word-mcp: <tool>") started
 → yield (app, doc, guard) to the tool body            [all edits via Range; Selection banned]
 → finally: end undo record (IsRecordingCustomRecord-guarded)
            guard.restore() → state_restore_failed[] into result
            doc = app = None; CoUninitialize
```

**Per-call flow inside a tool:** validate args → session (above) → locate content by Range/paragraph/bookmark/Find with inflated-range fallback → mutate through `guard.set` for any app/doc state → chunk >30KB text, reject `\x07` → build result with the standard fields: `undo_grouped`, `document_dirty` (`not doc.Saved`), `autosave_on`, `state_restore_failed`, plus tool payload. `doc.Save()` only when `save=True`.

**State guard:** LIFO snapshot-on-mutate (`StateGuard`). Always-restore: `TrackRevisions`, `UserName`/`UserInitials`, any `app.Options.*` touched, `ScreenUpdating` (only toggled on >20-mutation batches, with daemon-thread repair retry). Never-write: `Visible`, `Saved`, `AutoSaveOn`, window/view state, clipboard, `DisplayAlerts` (live instance only).

**Error taxonomy (typed, user-actionable):** `WordNotRunning` · `DocumentNotOpenInWord` · `ProtectedViewRefused` · `WordBusy` (retryable; dialog/Backstage) · `WordBlocked` (probe timeout; long sync op) · `WordDisconnected` (Word closed mid-call; warn "edit may be partially applied — Ctrl+Z undoes the partial step"). Recovery tool `word_live_repair()`: fresh attach; `ScreenUpdating=True`; `DisplayAlerts=wdAlertsAll`; end orphaned undo records.

**Pre-ship local empirical tests (scripts specified in Topics 2, 3, 7):** (1) UndoRecord empty-stack + wdReplaceAll crash probe; (2) ScreenUpdating survival after client hard-kill; (3) COM call behavior during Korean IME composition. Each result goes to BUILD_LOG and gates the corresponding default (replace_all grouping, ScreenUpdating watchdog paranoia level, IME retry budget).
