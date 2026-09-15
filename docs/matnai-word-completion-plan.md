# MatnAi: Word-only completion implementation plan

Status: offline MVP implemented; see [implementation and verification status](matnai-mvp.md)
for completed work and outstanding gates. Baseline recovery commit: `0048ba9`.
Recovery tag: `checkpoint/pre-matnai-2026-09-15` (local only).

## Scope and ribbon

One existing VSTO add-in, one installer, offline operation, no desktop companion,
cloud model, browser integration or sentence generation. No second dictionary or
replacement spelling engine. Context means modest ranking of word completions,
not a claim to predict fluent sentences. Personal learning is optional and local.

Updated UI decision: use one **MatnAi** group on the existing **Ўзбек Орфо** tab,
after corrections and before the dictionary. The original separate-tab proposal was
replaced at the user's request. Existing commands remain intact. The functional
categories below share that single ribbon group.

| Group | Controls |
| --- | --- |
| So'z takliflari | Ёқиш toggle; ready/loading/paused status |
| Moslashtirish | Sozlamalar; Shaxsiy o'rganish toggle (off until explicit consent) |
| Yordam | Qisqa qo'llanma; local diagnostics export |

Settings: script Auto/Latin/Cyrillic, minimum prefix length (default two letters),
result count (default three), configurable conflict-checked accept shortcut,
personal-learning consent and reset. Reset removes only prediction-learning data
after confirmation, never the main or custom dictionary. Labels are provisional.

## Module boundaries within the current project

| Responsibility | Proposed owner and reuse |
| --- | --- |
| Lexicon, spelling, morphology, transliteration | Existing DictionaryService, UzbekMorphAnalyzer, suffix JSON, TextHelper, TransliterationService |
| Completion request/results | CompletionRequest and CompletionCandidate; plain immutable data, no Word COM objects |
| Retrieval and ranking | WordCompletionEngine and CompletionIndex; document-independent, unit-testable |
| Suffix candidates | Bounded MorphologyCompletionProvider consuming the existing rule set and validating via the existing analyzer |
| Word lifecycle and scheduling | WordCompletionController owned/disposed by AddInRuntime |
| Read/insert safety | WordCompletionContextReader and CompletionInsertionService |
| UI | Non-activating CompletionPopup plus existing ribbon partial/designer integration |
| Optional learning | LocalCompletionPreferences with versioned, atomic storage and explicit consent |

Inspect all consumers before adapting the existing IPredictionEngine interface: it
currently describes n-gram learning and lacks cancellation and stale-result identity.
Reuse suitable settings (PredictionsEnabled, MaxPredictions, MinPredictionLength,
PreferredScript), not parallel settings files. Existing AutoLearn defaults true;
that legacy value is NOT consent. Introduce an explicit versioned consent state,
defaulting off for both new and upgraded installations. Preserve unrelated settings.

## Fast request path

1. While enabled and Word's editing window is active, detect caret/prefix changes.
   Prototype selection events plus a bounded UI-thread timer; measure missed events
   and polling cost before choosing the final cadence. Do not install a system-wide hook.
2. On Word's UI thread, read only the current token and bounded preceding context
   (initial cap 256 characters), with document/window/story identity and request ID.
   Do not scan the document or pass COM ranges into background work.
3. Debounce/coalesce requests (initial target 40-60 ms); cancel obsolete requests.
   Query an immutable in-memory prefix index on a worker. No per-keystroke disk I/O,
   full-dictionary fuzzy search, or full suffix enumeration.
4. Rank a bounded candidate set, return top three, then recheck identity/caret/text
   on the UI thread. Discard stale results even if cancellation arrived too late.
5. Position the popup near the caret without taking typing focus. Hide it when
   focus, document, selection or visibility changes; suppress same-request redisplay
   after Escape until the prefix changes or the user explicitly requests suggestions.

Build script-aware sorted prefix indexes off-thread from a dictionary snapshot;
binary-search prefix ranges, cache bounded top candidates and atomically publish ready
indexes. Maintain a small custom-word overlay and generation-based cache invalidation.
Respect apostrophe normalization, capitalization and source script; do not blindly
transliterate proper names or include multiword entries in a single-token completion.
Loading must leave Word usable, with prediction temporarily unavailable.

Suffix completion is generation, whereas today's parser primarily analyzes forms.
Implement only conservative, bounded generation using the same JSON constraints,
POS, allomorph and root-flag checks; run generated candidates through the analyzer.
Index imported standalone endings separately and never recursively combine them.
Do not expand every root by all 2,333 imported endings. Cap root/candidate work and
return fewer suggestions if the budget is exceeded. Existing acceptance does not
prove linguistic correctness; imported-rule review remains a release dependency.

Initial ranking: prefix match, supported morphology, script/case consistency, available
frequency evidence, then deterministic ties. Confirm the frequency seed actually exists
and is wired before depending on it. Optional local acceptance counts and bounded
previous-token associations may later adjust ranking; no invented frequency data.

## Interaction and insertion safety

Start with ordinary body-text typing at a collapsed caret at the end of a token.
Suppress predictions in protected/read-only content, fields, unsupported content
controls, non-body stories, selected text, dialogs and active input composition.
Tables, tracked changes and mid-token editing stay disabled until tested explicitly.

Use click or a configurable shortcut to accept; prototype a conflict-checked shortcut
such as Ctrl+Space before deciding the default. Do not hijack Tab, Enter or Space.
Escape dismisses. Native shortcuts keep working when the popup is hidden or stale.
Keyboard navigation and screen-reader behavior are prototype acceptance criteria.

Before insertion, verify the same active document/window/story, collapsed selection,
expected caret position and exact original token text. Insert only the missing tail
for an exact surface prefix; if normalization requires replacement, replace only the
verified token span and preserve formatting. Never use blind keystroke injection.
Coordinate with AutoCorrectService to avoid reentrant double edits. Each acceptance
must undo as one operation and leave surrounding text/formatting unchanged.

Prototype positioning with Word's [Window.GetPoint](https://learn.microsoft.com/en-us/office/vba/api/word.window.getpoint);
Microsoft documents an error for a non-visible target. Fail closed: hide the popup,
do not scroll the document to force a location. Prototype custom undo grouping with
[UndoRecord](https://learn.microsoft.com/en-us/office/vba/word/concepts/working-with-word/working-with-the-undorecord-object),
and test it rather than assuming ordinary insertions have the desired undo behavior.

## Milestones and acceptance gates

1. **Interaction spike:** hard-coded suggestions only; test typing detection, popup,
   focus, shortcut conflicts, stale rejection and undo in real Word. No learning.
   Stop for a design decision if safe positioning/input capture cannot be demonstrated.
2. **Core completion:** implement prefix index, Latin/Cyrillic normalization, custom-word
   updates, deterministic ranking and cancellation. Test against the full 216,339-word
   baseline and rapid scripted prefix changes, not just repeated known words.
3. **Bounded morphology:** add prefix-aware suffix candidates and regression cases
   for valid/invalid ordering, allomorphs, flags and standalone imported endings.
4. **MatnAi UI integration:** add tab/groups/settings; wire runtime ownership,
   focus/document lifecycle, AutoCorrect coordination and safe enable/disable cleanup.
5. **Optional personal ranking:** explicit consent, acceptance-only word/count learning,
   capped local storage, clear/reset and corruption recovery. No automatic document
   ingestion, raw document logging or network transmission.
6. **Validation/pilot:** unit and actual Word tests, opt-in usability pilot, large-document
   latency runs and the existing six-target Office installation/upgrade matrix.

Proposed performance budgets, NOT measured promises: warm engine p95 <= 15 ms;
UI snapshot/render work p95 <= 5 ms per operation; last keystroke to visible suggestion
p95 <= 100 ms including detection/debounce. Measure these separately on a recorded
reference machine, in small and 50,000+ token varied documents, cold/warm states,
both scripts and rapid typing. Record p99, startup time and incremental memory too.
Tune polling/debounce from end-to-end measurements rather than adding independent
timer delays that exceed the latency budget. No busy polling when disabled/unfocused.

Correctness gates: zero unintended edits and zero stale-result insertions in automated
scenarios; acceptance-only mutation; single-step undo; no regressions to existing
spelling, transliteration, dictionary edits or user data. Test switching/closing windows,
scroll/zoom, multi-monitor DPI, deletion/paste, IME, popup clicks, protected documents,
tables and tracked changes (including suppression of unsupported contexts).

Pilot metrics: useful top-three completions, acceptance rate, net characters saved,
dismissals, immediate undos and reported accidental insertions. An undo is a proxy,
not proof of an accidental insertion. Diagnostic collection is separately opt-in;
export aggregate counts/timings, not words or document contents. Agree quality targets
before the pilot; no production quality claim from the engineering seed corpus.

## Git and rollout

Baseline checkpoint already exists locally; it is a recovery snapshot, not a certified
release. Ignored build outputs, personal settings, certificates and TestResults are
not in Git. No remote backup has been made.

Implement later on `feature/matnai-word-completion` with separate commits per milestone
and feature disabled by default. Preserve the checkpoint tag. Prefer reverting feature
commits for rollback. To inspect the baseline without overwriting current work, create
a separate worktree from the tag; do not use destructive reset commands.

This document records the original plan; actual implemented scope is tracked in the MVP
status document. Desktop expansion is explicitly out of scope.
