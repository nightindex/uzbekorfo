# MatnAI Word completion MVP

The original lexical MVP is now extended by [document-based phrase prediction](prediction/README.md). Current corpus evidence and unfinished production gates are in [validation](prediction/validation.md).

Implemented on `feature/matnai-word-completion`. The pre-feature recovery tag is
`checkpoint/pre-matnai-2026-09-15` (commit `0048ba9`). No separate application or
new dictionary/morphology runtime was introduced.

## Using it

Build/run the add-in in Word through the existing Visual Studio/VSTO workflow.
The existing **Uzbek Orfo** tab contains a single **MatnAI** group, after the correction
group and before the dictionary group. It contains three primary controls:

- **Completion:** MatnAI offers suggestions automatically after the user types the
  configured minimum word prefix. The dedicated toggle enables or disables it.
  There is no manual suggestion button because clicking a Ribbon button removes
  focus from the Word editing surface.
- **Settings:** minimum prefix length and result count; optional personal learning,
  confirmed reset, and a manual index refresh for troubleshooting. Dictionary saves
  refresh the index automatically.
- **Help:** keyboard and scope instructions.

The group is declared in `UzbekOrfoRibbon.Designer.cs`'s `InitializeComponent`,
so MatnAI is available in the Visual Studio Ribbon Designer as well as at runtime.
Open `UzbekOrfoRibbon.cs` with View Designer and find **MatnAI** on the **Uzbek Orfo** tab. Handler
code remains in the nested code-only `UzbekOrfoRibbon.MatnAi.cs` file. The static
`eng/test-ribbon-layout.ps1` guard runs as part of UI checks; it is not an IDE test.

Type at least two letters in ordinary document body text. MatnAI shows only the
untyped continuation as gray inline text. Press Tab to accept it, Down to open the
compact alternatives list, and then Up/Down to select. Escape dismisses. The existing
Ctrl+Alt+Up/Down and Ctrl+Alt+Right shortcuts remain available. These keys are captured
only while a current suggestion is visible and Word's document editing surface has
focus; otherwise Word receives them normally. Enter and Space are never captured.
Ctrl+Z reverses an accepted completion in one step.

Automatic completion defaults off. Personal learning requires a separate explicit
opt-in, even if the historical AutoLearn setting was true. The opt-in store contains
only accepted words/phrases and counts in the encrypted versioned store under
`%LocalAppData%/UzbekOrfo/MatnAI`. Older accepted-word preferences migrate when
learning is enabled. Ordinary typing is not recorded or transmitted. Turning
learning off stops collection and use while retaining preferences. Reset clears
acceptance preferences and legacy recovery copies, not document collections or dictionaries.

## Implementation boundaries

- Immutable prefix index is built from a snapshot of the existing DictionaryService;
  dictionary saves trigger refresh. Queries do bounded indexed retrieval, not full
  dictionary fuzzy scans. Superseded builds/queries are cancelled.
- Latin/Cyrillic script, typed casing and apostrophe forms are preserved.
- The worker proposes bounded suffix completions from the existing suffix JSON.
  On Word's UI thread, the existing dictionary and morphology analyzer validate all
  proposals before display. No recursively generated suffix combinations are added.
  Conservative generation requires an exact known stem and a partially typed ending;
  it does not guarantee completion of every valid inflection or mutated stem.
- The displayed word stays first while the user continues typing that same word.
  New predictions honor optional acceptance counts, then favor continuations of two
  or more letters before shorter completions. A remaining single letter can still
  complete the displayed word. The new document-collection engine also ranks observed
  contextual phrases; it is statistical, not neural sentence generation.
- A click-through layered ghost window follows the caret without modifying the document.
  Down opens a compact, non-activating alternatives popup with accessible list items and
  stale-action rejection. Both overlays use the Word window's DPI context and the caret's
  monitor work area. Document identity, window, token span and exact typed prefix are
  rechecked immediately before insertion. Only the missing tail is inserted.
- Ghost text follows the native caret and the last typed character's font at Word's
  current zoom. Its font ascent/descent determines the baseline. When the native caret
  is unavailable, the fallback removes paragraph spacing from Word's range bounds.
- Word COM access remains on the UI thread. Background queries receive plain strings
  and snapshots. Completion temporarily suspends the existing autocorrect service
  during insertion to prevent reentrant corrections.
- Protected/read-only documents, selected text, tables, tracked changes, fields,
  non-body stories, active IME composition and documents containing content controls
  are conservatively excluded. The content-control exclusion currently applies to
  the whole document, not just the current control.

## Verification

Run `eng/test-unit.ps1 -Configuration Debug`, `eng/test-word.ps1 -Configuration Debug`
and `eng/test-ui.ps1 -Configuration Debug` after a Debug build.
The opt-in `eng/test-word.ps1 -Configuration Debug -VisualGhost` briefly opens its own
temporary Word document and compares cropped ghost text with accepted Word text. It
checks both native-caret and range-fallback placement in Arial, Calibri and Times New
Roman, Latin/Cyrillic, at 100%/150% zoom (24 cases). Captures are saved in the harness
output folder. It does not inspect existing documents or change add-in settings.

Current measured evidence:

- 69 unit tests passed, including stable whole-word continuation, configurable ghost-tail eligibility, cancellation, optional preference ranking, settings
  consent migration, script/apostrophe preservation and deferred validation.
- Real-Word smoke passed: accepted tail insertion, single Undo, stale prefix rejection,
  other-document rejection, selected-text/tracked-change suppression, suffix validation,
  repeated index rebuild and idempotent disposal. Tests use isolated settings directories.
- The completion UI tests passed layered ghost rendering, focus preservation, accessible
  continuation/list behavior, compact popup sizing, stale-action rejection, and placement
  on negative-coordinate and differently scaled monitor geometries. The full UI suite
  passed from 100% through 400% scaling.
- Visual Word checks passed all 24 cases; ghost and accepted glyph top/bottom positions
  differed by at most one screen pixel in this run. This does not replace testing on
  the user's actual pair of monitors.
- The latest local Word smoke run built the full lexical index in 168 ms. Across 1,200
  warmed queries, lexical engine p95 was 0.046 ms. These are **not** keystroke-to-suggestion measurements
  and do not include suffix validation, Word polling, UI rendering or cold startup.

## Remaining work before a production claim

Measure actual typing-to-suggestion p95/p99 on varied large documents; test the integrated
ribbon and overlays with actual typing, zoom, IME, mixed-DPI monitors and Narrator. The
automated overlay test is not a full assistive-technology certification. Validate keyboard
shortcut conflicts on target machines. The 100 ms end-to-end target remains unproven.

Contextual ranking, configurable acceptance shortcuts, script override, aggregate
diagnostic export and a real-user usefulness/accidental-insertion pilot remain planned
work. Existing independent linguistic review and the six-target clean installation/
upgrade matrix remain pending. No signed installer was published or installed here.
