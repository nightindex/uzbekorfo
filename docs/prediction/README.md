# MatnAI document prediction

Status: implemented experimental offline phrase prediction; not a production-quality or legal-accuracy certification. This extends the original [word-completion MVP](../matnai-mvp.md). Recovery checkpoint: `checkpoint/pre-document-prediction-20260916`.

## Using collections

1. Open Word's existing MatnAI group, then **Созламалар → Ҳужжатлардан ўрганиш**.
2. Choose **Янги тўплам**, enter a name, and use **Файл қўшиш**, **Папка қўшиш**, or drop files/folders onto the page. Set the subfolder checkbox before adding a folder.
3. Review the discovered files and press **Ўрганиш**. Progress and per-file outcomes appear on the page. **Тўхтатиш** cancels before publication; a completed atomic publication is already committed.
4. Check the private collection to use it in the current document. Importing alone does not activate it. Other open documents have independent selections; closing a document clears its selection.
5. Enable suggestions on the ribbon. Type two letters for current-word completion, or finish a word and type a space for contextual suggestions. Tab accepts exactly the visible gray continuation, Esc dismisses, and Down opens alternatives. Existing Ctrl+Alt shortcuts remain. One Undo reverses acceptance.

The built-in legal collection is included automatically when prediction is enabled. Rename, refresh, source removal, and collection deletion apply immediately, independently of the settings dialog's Save/Cancel buttons. Removal and deletion require confirmation and never delete original documents. Manually removed source paths remain excluded from folder refresh; explicitly adding the file again restores it. Missing files during refresh retain their previous contributions until explicitly removed.

The **Шахсий ўрганиш** page controls a separate opt-in: accepted words/phrases and counts only. Ordinary typing is not recorded. Clearing acceptance learning leaves document collections and the spelling dictionary unchanged. Removing a source clears acceptance preferences for its collection conservatively; deleting a collection removes its acceptance preferences.

## Implementation

Both Cyrillic and Latin legal models are now embedded and automatically enabled.
The engine filters candidates by the current text's script. To rebuild only the
Latin model, run `./eng/build-prediction-corpus.ps1 -Script latin`; its sources are
in `corpus/raw-latin/`, and its separate reports are in `docs/prediction/latin/`.
See [Latin model build and measurements](../legal-latin-corpus.md).

- `Prediction/DocumentExtractor.cs`: managed DOCX ZIP/XML, binary Word 97–2003 via HWPF, HTML-as-DOC via HtmlAgilityPack, and TXT extraction. No Word automation, macros, scripts, or external-resource fetching during import. Comments, deleted revisions, headers/footers and known LexUZ navigation/metadata are excluded where supported.
- `CollectionImporter`: one cancellable import job, SHA-256 file/passage deduplication, paragraph/sentence boundaries, per-source contributions, and unigram through seven-token observed sequences. Seven tokens represent up to two context words plus five continuation words. Deduplicated copies can regain ownership when the original source is removed.
- `CollectionStore`: versioned compressed JSON; private payloads use Windows CurrentUser DPAPI. Staging is encrypted, publication is atomic, and valid backups permit recovery. Explicit source removal discards the previous backup; collection deletion removes its primary, backup, and recognized interrupted staging files.
- `PhrasePredictionEngine`: immutable indexed contexts, shorter-context backoff, script filtering, repeated-support gating, bounded collection-scoped acceptance boost, and up to five-word candidates. Matching uses normalized forms and preserves the exact typed prefix; imported display spellings are retained in source metadata, but arbitrary source capitalization is not yet fully reproduced by phrase output.
- `WordCompletionController`: snapshots at most 512 preceding characters on Word's UI thread, sends plain requests to workers, rejects stale results after context/selection/collection changes, and combines observed phrases with existing dictionary/morphology completion. Imports never change spelling correctness or authorize suffix combinations.
- Ghost rendering fits complete words within the supported text area. Acceptance rechecks surrounding text and inserts only the visible continuation in a single custom Undo record. Tables, tracked changes, protected documents, fields, content controls, non-body stories and other existing unsupported contexts remain excluded. Multi-column inline placement is conservatively suppressed.

This is local statistical learning, not a neural network and not a reproduction of another product's proprietary algorithm. There is no cloud service, GPU requirement, folder watcher, PDF parser, or automatic script conversion. Cyrillic is preserved as Cyrillic, not declared to be Uzbek by script detection alone.

Ghost stability/casing fix: a visible completion is retained while exact matching
typing consumes it, including across spaces; incompatible edits still invalidate
it. Unchanged and width-clipped tails reuse the layered surface rather than
redrawing on every polling tick. Background sampling is outside the ghost glyphs.
`Data/prediction_casing.json` supplies display-only casing for **Ўзбекистон Республикаси**
and **O'zbekiston Respublikasi**, including observed endings on the last word.
This applies to phrase suggestions and dictionary fallback without retraining or
changing spelling validity. Text already typed is never silently recased.

The layered window must be shown before publishing its pixel surface: WinForms
can reset a surface drawn before `Show`. Fallback positioning uses Word's character
top without an extra scaled vertical offset. The opt-in Word visual harness now
compares ghost/accepted glyph bounds within one pixel and requires an unlocked,
visible desktop; a successful nonvisual test is not proof of visual alignment.
An owned ghost window could also remain behind Word. After publishing its surface,
the renderer now raises it without activation, only when Word is foreground, and
without making it globally topmost. All 24 visual font/script/zoom/anchor cases
passed on the tested 96-DPI Word window; see [pilot evidence and limits](pilot.md).

## Storage and limits

Private data lives in `%LocalAppData%/UzbekOrfo/MatnAI`:

- `<id>.private.gz.dpapi`: source references, hashes, derived vocabulary and phrase counts; no copied source documents. Recovery copies use `.bak`.
- `acceptances-v2.bin`: encrypted collection-scoped accepted words/phrases and counts. The previous `%AppData%/UzbekOrfo/matnai_acceptances.tsv` preferences migrate when learning is enabled; reset removes that legacy file and its backup to prevent reimport.
- `metrics-v1.tsv`: optional local session aggregate counts and latency buckets, with no document text. Key-to-ghost buckets include the polling delay for observed editing keys; mouse, menu and some IME edits have no such sample. Nothing is transmitted. Keyboard-based immediate-Undo tracking covers Ctrl+Z within five seconds when the previous context is restored; it does not cover every Ribbon Undo path.

Derived phrases can contain confidential information: encryption does not anonymize them. Import only documents the user is authorized to process. Windows user-bound encryption is not portable to another account/machine by copying the payload alone.

Initial limits: 50 MiB/file, 1,000 files and 500 MiB/job; bounded extracted text/XML and 200,000 stored model records per collection. Rare sequences may be compacted to stay within limits, with a message in the import report. These are safety limits, not guarantees that every format variant is readable or that an unlimited number of loaded collections fits the memory target. Password-protected, corrupt, inaccessible and unsupported files are reported rather than opened through Word.

## Reproduce and validate

See the [selective-suggestion pilot](pilot.md) for the current after-space
policy, comparison reports, local counters, and remaining human/Word checks.
The original corpus reports predate this policy; use `pilot-report.json` in
each script's report directory for the controlled comparison.
The latest [memory comparison](memory-comparison.json) records snapshot-local
sharing and compact indexes; collection settings retains built-in metadata only.
This reduces retained memory without changing the replay metrics, but does not
certify peak memory or the reference-PC target.

Build the add-in with the repository's normal Visual Studio/VSTO build, then run:

```powershell
./eng/test-unit.ps1 -Configuration Debug
./eng/test-ui.ps1 -Configuration Debug
./eng/test-prediction-ui.ps1 -Configuration Debug
./eng/test-word.ps1 -Configuration Debug
```

`eng/build-prediction-corpus.ps1` rebuilds the embedded legal collection and evaluation reports from the downloaded `corpus/raw_cyrillic/**/*.doc` files (falling back to legacy `corpus/raw/`). It requires the .NET SDK, Visual Studio compiler, restored packages and the local raw corpus. Raw source documents remain ignored by Git. The script replaces generated reports, including the unreviewed context template: preserve a completed human-review record separately before rerunning it. `-PilotEvaluation` writes only the pilot report, leaving models and review templates unchanged.

See [validation evidence and remaining release gates](validation.md), [machine-readable corpus report](corpus-report.json), and the [200-context review template](review-contexts.json). The model artifact is built from all public sources only after evaluating a separate training-only model on held-out groups; reported test scores must not be measured against the final all-source artifact.

## Native Word prediction coexistence

MatnAI does not change Word's global prediction settings. Two enabled inline systems can show competing suggestions or intercept Tab differently. For a controlled pilot, test each separately first, then both together; users can disable MatnAI with its ribbon toggle. If suggestions are misplaced, record Word build/architecture, monitor scale, zoom and editing context, then test moving the same window between monitors. Never treat a predicted legal phrase as confirmation that a law is current or legally appropriate.
