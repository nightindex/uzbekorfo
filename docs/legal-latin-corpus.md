# Uzbek Latin legal prediction corpus

The Cyrillic legal corpus remains under `corpus/raw/`. Uzbek Latin exports are
stored separately under `corpus/raw-latin/`, which prevents a corpus rebuild from
silently mixing scripts.

Run the resumable downloader from the repository root:

```powershell
./eng/download-legal-latin-corpus.ps1
```

It reads the checked-in 57-source provenance list from
`docs/prediction/corpus-report.json`, then downloads each corresponding official
Latin export from `https://lex.uz/docs/-{id}?type=doc`. The leading minus selects
the Uzbek Latin publication. For example:

<https://lex.uz/docs/-20596?type=doc>

The downloader validates that every response is a Word-compatible HTML document
whose text is predominantly Latin script, writes each file atomically, and records
the exact source URL in `corpus/raw-latin/manifest.json`. Existing valid downloads
are retained; use `-Force` only to replace them with a fresh download.

## Build and use the Latin model

```powershell
./eng/build-prediction-corpus.ps1 -Script latin
```

This extracts the Latin files, verifies their script and manifest provenance,
groups duplicate/related documents before splitting, evaluates a training-only
model on held-out documents, and publishes the final all-source artifact at
`src/UzbekOrfoAddIn/Data/legal_prediction_latin.collection.gz`.
Extraction failures retain the previous usable artifact.

Rebuild the add-in in Visual Studio after regenerating the artifact. The DLL
embeds both legal collections. MatnAI automatically includes both when enabled;
the existing predictor chooses matching-script candidates from the typed prefix
or, after a space, the preceding word. Settings lists the Latin collection as
**Ҳуқуқий ҳужжатлар (лотин)**. No manual collection selection is necessary for it.

The default build command still builds the Cyrillic artifact. Latin evaluation
and the 200-context review template are stored under `docs/prediction/latin/`,
separately from the Cyrillic reports. The manifest retains negative Lex.uz URLs.

The 2026-09-17 Latin run extracted all 57 files: 947,919 words after within-source
passage deduplication, 5,132 repeated normalized sentences across sources, and
120,561 retained sequence records. The compressed artifact is 629,958 bytes.
On 1,500 held-out replay contexts, the deployed thresholds produced 23.9% top-1
and 33.3% top-3 next-word accuracy, 87.3% coverage, and 14.3% simulated keystrokes
saved. Warm query p95 was 1.38 ms (4.09 ms during import) on this 16-processor,
16-GB development machine. The Latin engine index alone used about 39.6 MiB;
this is not total Word/add-in memory or a combined-collection reference-PC test.

The report's `selectedThresholds` are advisory validation results;
`evaluatedRuntimeThresholds` records the existing shared runtime settings
(0.4 probability, 0.0 margin) actually used for held-out scores. Text comparison
normalizes case and apostrophe variants. Independent linguistic review and the
supported Office installation/pilot matrix remain pending.
