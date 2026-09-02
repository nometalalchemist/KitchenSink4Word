# KS4W v2 Phase 6b: Discoverability Round Log

Round start: 2026-09-03 03:12 KST. Branch v2-dev @ 574e7b4.
Method: in-process fastmcp Client session against word_mcp.server in
LITE mode (28 tools; global Visibility transform exactly as main()
wires it), driven over a local TCP harness. Each scenario is a FRESH
subagent (Sonnet, general-purpose, zero project context) whose prompt
contains only the live lite tool listing (names, params, full
docstrings), the task, and the JSON call protocol. The orchestrator
relays each call to the live session and returns the raw wire result
verbatim; after an enable, the subagent receives the real new surface
(delta listing). Max 10 turns per scenario.

PASS = the agent reaches the correct pack via enable_tools (through
the enable_tools menu, get_workflows, or a signpost refusal) and then
calls a correct pack tool for the task. Session reset (fresh lite
surface) between scenarios.

Note: the intended Fable-powered subagents were flagged by the model's
safeguards on the scenario prompt (large embedded tool listing plus
strict JSON-only protocol); the round runs on Sonnet, which is the
harsher discoverability test anyway (a median agent, not a frontier
one).

---
## Scenario 1: references — PASS

Task: "Add a citation to Smith (2020) after the second paragraph of
chapter.docx and insert a bibliography at the end."

| turn | call | outcome |
|---|---|---|
| 1 | get_document_view(chapter.docx) | orientation, 5 blocks |
| 2 | enable_tools(packs=['references']) | ok, +8 tools, list_changed |
| 3 | get_workflows(task='format-citations') | recipe consulted |
| 4 | detect_citation_system | system=none (followed recipe) |
| 5 | enable_tools(packs=['academic']) | for insert_reference_list |
| 6 | manage_source(action='add', tag='Smith2020', Smith 2020) | ok |
| 7 | insert_citation(anchor 25e3, position='end') | BAD_PARAMS: only 'after' supported |
| 8 | insert_citation(anchor 25e3, position='after') | ok, (Smith, 2020) placed |
| 9 | insert_reference_list(type='bibliography') | ok, Bibliography at end |
| 10 | get_document_view | verified, then {"done": true} |

Discovery route: MENU (enable_tools docstring), on the very first
action turn after orientation. The agent also found get_workflows
unprompted and followed the format-citations recipe, and correctly
enabled academic as the second pack because insert_reference_list
(bibliography) lives there; the manage_source docstring pointed it
there ("render the list with insert_reference_list(type='bibliography')").
Verdict: PASS (pack via menu, turn 2; correct pack tool manage_source
turn 6 / insert_citation turn 8). Task itself completed end to end.
Note: one BAD_PARAMS on position='end' self-corrected next turn from
the refusal message; refusal copy did its job.

## Scenario 2: review — PASS

Task: "Accept all tracked changes in draft.docx and delete every
comment." Fixture: 1 tracked replacement (2 revision elements) + 2
comments.

| turn | call | outcome |
|---|---|---|
| 1 | accept_revisions (v1 name) | "Unknown tool: 'accept_revisions'" (no signpost; name not registered in v2) |
| 2 | mcp__word__accept_revisions (env-contaminated v1 name) | "Unknown tool" again |
| 3 | enable_tools(packs=['review']) | ok, +9 tools, list_changed |
| 4 | resolve_revisions(action='accept') | ok, 2 revisions resolved |
| 5 | get_comments | 2 comments listed |
| 6 | manage_comment(action='delete', id 0) | ok |
| 7 | manage_comment(action='delete', id 1) | ok |
| 8 | get_comments | [] verified |
| 9 | get_tracked_changes | [] verified, then {"done": true} |

Discovery route: MENU, self-corrected after two unknown-tool errors.
The agent first guessed the v1 tool name accept_revisions (turn 2's
mcp__word__ prefix shows the guess came from v1.6 tools visible in the
scenario agent's OWN environment, a harness contamination the round
now guards against with an isolation line in every prompt; the
contamination made the test harder, not easier). Verdict: PASS (pack
via menu turn 3, correct pack tools resolve_revisions turn 4 and
manage_comment turns 6-7). Task completed and verified end to end.

FINDING (soft, no gate impact): a v1 tool name that no longer exists
in v2 returns a bare "Unknown tool: 'accept_revisions'" with no pack
hint, because the DisabledToolSignpost only covers registered-but-
disabled names. The V2_PLAN Section on option C/D predicted exactly
this hole. The agent recovered via the menu in one turn, so no RED;
logged as residual risk for the readout (candidate 2.1 improvement:
route known v1 names through migration/v1_to_v2.json to a signpost).

## Scenario 3: academic — PASS

Task: "Add a footnote to the third paragraph of thesis.docx saying
'See appendix for details.'" (Prompt now carries the environment-
isolation line; this and later agents made zero direct tool calls.)

| turn | call | outcome |
|---|---|---|
| 1 | get_document_view(scope 0-6) | orientation |
| 2 | get_document_view (repeat) | same view |
| 3 | enable_tools(packs=['academic']) | ok, +23 tools, list_changed |
| 4 | manage_note(action='insert', note_type='footnote', anchor 2056) | ok, footnote 1 added, then {"done": true} |

Discovery route: MENU (the "notes (foot/end)" phrase in the academic
menu line). Verdict: PASS (pack via menu turn 3, correct pack tool
manage_note turn 4, task completed).

## Scenario 4: assembly — PASS

Task: "Merge intro.docx and body.docx into one document called
combined.docx."

| turn | call | outcome |
|---|---|---|
| 1 | enable_tools(packs=['assembly']) | ok, +7 tools (menu route, before any read) |
| 2 | copy_document(intro.docx -> combined.docx) | ok |
| 3 | insert_document(target=combined, source=body) | ok, 3 paragraphs merged, then {"done": true} |

Discovery route: MENU ("multi-document work: insert/split documents"
line), enabled as the very first call. The copy-then-insert plan is
exactly the intended assembly idiom (copy_document's own lite
docstring adverts the pack: "Split/merge and multi-document work live
in the assembly pack"). Verdict: PASS. Task completed.

## Scenario 5: media-forms — PASS

Task: "Sort the table in report.docx by its second column,
descending."

| turn | call | outcome |
|---|---|---|
| 1 | enable_tools(packs=['media-forms']) | ok, +16 tools (menu route, first call) |
| 2 | get_table(table_index=0) | read 5x3 table |
| 3 | sort_table(column=1, numeric=true, descending=true, has_header=true) | ok; verified order East 88 / North 42 / South 17 / West 3, then {"done": true} |

Discovery route: MENU ("table structure/styling (rows, columns,
merges, widths, sort)" line; the get_table lite docstring's
"Row/column surgery, styling, and sort: media-forms pack" advert
corroborated). Verdict: PASS. Correct table-power-tools-moved-to-
media-forms mapping confirmed (the Phase 4 relocation is
discoverable). Task completed with correct 0-based column index and
header pinned.

### Mid-round harness-integrity probe (coordinator challenge)

Challenge: did assembly/media-forms complete their pack calls without
enable_tools (which would mean lite was not enforced)? Transcript
answer: NO. Both agents called enable_tools as their FIRST action
(assembly turn 1, media-forms turn 1; both results carried
list_changed and the +7/+16 tool bills), and only then called
insert_document / sort_table. Probe on a fresh reset session, before
scenario 6 ran:

- visible tool count at session start: 28 (exact lite surface)
- sort_table visible: False; insert_document visible: False
- sort_table call in lite: signpost refusal naming media-forms and
  the exact enable_tools call
- insert_document call in lite: signpost refusal naming assembly

Every scenario in this round started from a reset session with this
same 28-tool surface (reset confirmed before each spawn). Lite
enforcement and Phase 4 pack membership are real at the wire.

## Scenario 6: com-live — PASS

Task: "Export report.docx to PDF."

| turn | call | outcome |
|---|---|---|
| 1 | enable_tools(packs=['com-live']) | ok, +13 tools (menu route, first call) |
| 2 | com_export_pdf(report.docx) | ok, report.pdf written (112,096 bytes), then {"done": true} |

Discovery route: MENU ("drives the Word app: PDF import/export"
line). The COM call itself ran for real (no user Word instance was
open; verified before the call), the invisible instance exported and
exited, and the zombie check after the run found no WINWORD.EXE.
Verdict: PASS on both the gate criterion (correct pack + correct
tool) and the runtime result.

## Scenario 7: protection-io — PASS

Task: "Add a DRAFT watermark to manuscript.docx."

| turn | call | outcome |
|---|---|---|
| 1 | enable_tools(packs=['protection-io']) | ok, +6 tools (menu route, first call) |
| 2 | set_watermark({text: 'DRAFT', color: 'gray', diagonal: true}) | ok, 1 header part, then {"done": true} |

Discovery route: MENU ("document protection, watermarks, redaction"
line). Verdict: PASS. Task completed.

## Control 1: lite-only replace — PASS (clean)

Task: "Replace the word 'foo' with 'bar' everywhere in notes.docx."

| turn | call | outcome |
|---|---|---|
| 1 | search_and_replace([{find: 'foo', replace: 'bar'}], scope='all') | ok, 4 replaced |
| 2 | {"done": true} | |

No pack enabled, no get_workflows detour. The lite copy did not
oversell packs for a plain replace.

## Control 2: lite-only structure — PASS (clean)

Task: "Show me the structure of this document." (outline_doc.docx)

| turn | call | outcome |
|---|---|---|
| 1 | get_document_view(detail='structure') | ok, 3-heading outline |
| 2 | {"done": true} | correct structure summary |

No pack enabled. (get_outline was the other acceptable lite answer;
the agent chose the view tool's structure detail, equally correct.)

---
## Round summary

| # | pack | discovery route | turns to enable | pack tool called | verdict |
|---|---|---|---|---|---|
| 1 | references | menu (+ get_workflows consulted after) | 2 | manage_source, insert_citation (+academic: insert_reference_list) | PASS |
| 2 | review | menu, after 2 unknown-tool errors on v1 names | 3 | resolve_revisions, manage_comment | PASS |
| 3 | academic | menu | 3 | manage_note | PASS |
| 4 | assembly | menu | 1 | insert_document (after copy_document from lite) | PASS |
| 5 | media-forms | menu | 1 | sort_table | PASS |
| 6 | com-live | menu | 1 | com_export_pdf (ran for real; PDF produced; no zombie WINWORD) | PASS |
| 7 | protection-io | menu | 1 | set_watermark | PASS |
| C1 | (control) | none needed | n/a | search_and_replace (lite) | PASS clean |
| C2 | (control) | none needed | n/a | get_document_view (lite) | PASS clean |

7/7 scenarios PASS on first run each; 0 RED gates; 0 copy fixes
required; both controls clean (no pack overselling). Dominant
discovery route was the enable_tools MENU, which agents read straight
out of the lite listing; the signpost path was exercised only by the
mid-round integrity probe (sort_table/insert_document refusals
verbatim correct) and never needed by an agent, and get_workflows was
consulted once (scenario 1) and its recipe followed. Every scenario
started from a reset 28-tool lite session (verified before each
spawn; probe section above).

Round end: all seven packs reachable unprompted from lite by a fresh
Sonnet agent with nothing but the surface text. Gate: GREEN.

Suite after the round: 1186 passed, 1 skipped, 62 deselected (live)
in 253.54s (2026-09-03 03:52 KST). No live-relevant changes made; live
suite not required. No copy changes this round, so the docstring
budget and no-em-dash tests were exercised unchanged and stayed green.
