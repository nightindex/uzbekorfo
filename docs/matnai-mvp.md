# MatnAi Word completion MVP

Implemented on `feature/matnai-word-completion`. The pre-feature recovery tag is
`checkpoint/pre-matnai-2026-09-15` (commit `0048ba9`). No separate application or
new dictionary/morphology runtime was introduced.

## Using it

Build/run the add-in in Word through the existing Visual Studio/VSTO workflow.
The existing **Uzbek Orfo** tab contains a single **MatnAi** group, after the correction
group and before the dictionary group. It includes:

- **Completion:** Yoqish enables automatic completion; Takliflarni ko'rsatish
  requests a suggestion session without turning automatic completion on.
- **Settings:** minimum prefix length and result count; dictionary index refresh;
  optional personal learning and confirmed reset.
- **Help:** keyboard and scope instructions.

The group is declared in `UzbekOrfoRibbon.Designer.cs`'s `InitializeComponent`,
so MatnAi is available in the Visual Studio Ribbon Designer as well as at runtime.
Open `UzbekOrfoRibbon.cs` with View Designer and find **MatnAi** on the **Uzbek Orfo** tab. Handler
code remains in the nested code-only `UzbekOrfoRibbon.MatnAi.cs` file. The static
`eng/test-ribbon-layout.ps1` guard runs as part of UI checks; it is not an IDE test.

Type at least two letters in ordinary document body text. Click a suggestion to
accept it, or use Ctrl+Alt+Up/Down to choose and Ctrl+Alt+Right to accept.
Escape dismisses. These shortcuts apply only while a suggestion popup is visible
and Word's document editing surface has focus. Tab, Enter and plain arrow keys are
unchanged. Ctrl+Z reverses an accepted completion in one step.

Automatic completion defaults off. Personal learning requires a separate explicit
opt-in, even if the historical AutoLearn setting was true. The opt-in store contains
only accepted words and counts in `%AppData%/UzbekOrfo/matnai_acceptances.tsv`,
capped at 1,000 entries. No document/context text is collected or transmitted.
Turning learning off stops collection and use, but retains the file. Reset clears
the active model, not either dictionary; AtomicFile may retain a local .bak recovery
copy, as disclosed in the confirmation.

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
- Optional acceptance counts influence ranking. Without them, ranking is deterministic
  and favors shorter completions. There is no trained contextual or sentence model.
- A non-activating popup follows the caret, exposes accessible list items and rejects
  stale accessibility actions. Document identity, window, token span and exact typed
  prefix are rechecked immediately before insertion. Only the missing tail is inserted.
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

Current measured evidence:

- 59 unit tests passed, including cancellation, optional preference ranking, settings
  consent migration, script/apostrophe preservation and deferred validation.
- Real-Word smoke passed: accepted tail insertion, single Undo, stale prefix rejection,
  other-document rejection, selected-text/tracked-change suppression, suffix validation,
  repeated index rebuild and idempotent disposal. Tests use isolated settings directories.
- The popup-specific UI test passed focus preservation, accessible list/selection,
  explicit accessible acceptance and stale accessibility item rejection. The wider UI
  suite passed on standalone reruns. An earlier run failed the existing exceptions-grid
  row-count assertion at 300%; the cause is unconfirmed and was not silently repaired.
- One local run built the full lexical index in 169 ms. Across 1,200 warmed queries,
  lexical engine p95 was 0.023 ms. These are **not** keystroke-to-popup measurements
  and do not include suffix validation, Word polling, UI rendering or cold startup.

## Remaining work before a production claim

Measure actual typing-to-popup p95/p99 on varied large documents; test the integrated
ribbon and popup with actual typing, zoom, IME, multi-monitor DPI and Narrator. The
automated popup test is not a full assistive-technology certification. Validate keyboard
shortcut conflicts on target machines. The 100 ms end-to-end target remains unproven.

Contextual ranking, configurable acceptance shortcuts, script override, aggregate
diagnostic export and a real-user usefulness/accidental-insertion pilot remain planned
work. Existing independent linguistic review and the six-target clean installation/
upgrade matrix remain pending. No signed installer was published or installed here.
