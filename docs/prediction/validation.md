# Document prediction validation — 2026-09-16

Release status: experimental. Independent Uzbek legal-language review and the supported Office clean-install/upgrade matrix remain incomplete.

## Automated evidence

- Debug VSTO add-in build succeeded on .NET Framework 4.7.2.
- 90 unit tests passed, including managed extraction fixtures for TXT, DOCX, HTML-as-DOC and a genuine binary DOC; apostrophes and both scripts; duplicate ownership; missing refresh; corrupt input; cancellation; encrypted backup recovery; scoped learning; collection deletion and interrupted staging cleanup.
- The collection UI harness imports a real temporary folder, verifies that activation is opt-in per document, discovers a new file on manual refresh, and checks usable list widths/heights across repeated 96/144/192/96 DPI transitions. Screenshots are generated in its ignored output directory. Native Windows DPI context is kept consistent during the test.
- The broader UI harness covers existing dialogs, accessibility, ribbon structure and overlay geometry across 96–384 DPI. Automated scaling checks do not replace physical mixed-monitor testing.
- Real Word smoke checks passed on installed Word 16.0, build 16.0.17932, x64. Installed Click-to-Run product IDs include ProPlus2024Volume. These checks cover phrase-after-space acceptance, exact inserted text, single Undo, context-change rejection, per-document collection isolation and existing spelling/morphology exclusions.
- The final Word smoke run captured context in a 100,000-word document with p95 **12.92 ms**. This measures context capture, not end-to-end typing-to-ghost latency.

## Public-corpus experiment

The generated [report](corpus-report.json) records provenance URL, hash, family and split for every file. All **57/57** downloaded files extracted successfully: 945,771 word occurrences after within-source sentence deduplication, zero identical documents, and 5,384 cross-document duplicate sentence occurrences. The final artifact contains 120,334 retained sequence records and is 744,015 bytes. These counts are not a linguistically verified word inventory.

Source IDs, exact hashes, related civil-code parts and substantial sentence overlap group documents before the training/validation/test split. Sentences seen in training are excluded from validation/test replay. Version grouping is heuristic and still needs human leakage review. Validation selected probability 0.4 and margin 0.0, optimizing simulated saved keystrokes; runtime defaults use these frozen values.

On 1,500 held-out contexts mixing after-space and two-letter-prefix requests:

| Replay metric | Contextual model | Dictionary-only lexical baseline |
| --- | ---: | ---: |
| Exact next-word top-1 | 27.3% | 0.0% |
| Exact next-word top-3 | 36.0% | 0.33% |
| Suggestion coverage | 90.0% | 40.0% |
| Simulated remaining keystrokes saved | 16.3% | 0.0% |

Coverage means any suggestion was offered, not that it was correct. Keystroke savings require a complete offered continuation to match and subtract one acceptance key. The baseline is the dictionary prefix engine, without Word's runtime morphology validation/generation or user preferences. Its very low result on this mainly Cyrillic legal corpus is not evidence of general-purpose superiority. Separate after-space/prefix and Latin legal benchmarks remain desirable before release.

Measured on this development machine (16 logical processors, approximately 16 GB RAM): warm query p95 **1.57 ms**; query p95 during corpus import **7.00 ms**; additional managed phrase-index allocation approximately **39.5 MiB**. These are .NET 8 offline-engine measurements, not total Word-process memory or reference-PC acceptance tests. Index construction, collection deserialization, private models and concurrent import add memory beyond that figure. Local UI latency counters start when a changed context is detected, so they exclude up to one polling interval before detection.

## Release gates still open

1. Recruit an independent Uzbek legal-language reviewer. Review at least the 200 held-out contexts in `review-contexts.json`; all currently say `pending`. Record reviewer, source authorization, plausible alternatives, wrong or misleading completions, script/casing issues and decisions. Preserve signed review results separately from regenerable reports.
2. Expand representative legal writing beyond legislation, including permissioned drafts/contracts and both scripts. Audit related-version grouping and navigation extraction. Rerun comparison and freeze thresholds only after the reviewed validation set is accepted. Do not call these results spelling false-positive/false-negative measurements.
3. Run the pilot with opt-in local acceptance, dismissal and immediate-Undo counters. Investigate accidental insertions; no pilot user results have been fabricated.
4. Measure warm query p95 ≤50 ms, actual typing-to-ghost p95 ≤150 ms and total additional loaded prediction memory ≤100 MB on a four-core/8 GB/SSD reference PC, during import and in a 100,000-word document.
5. Complete the following matrix on clean machines/snapshots. For each row record OS, exact Office build/channel/architecture, installer version/hash, installation, upgrade with existing settings/collections, and uninstallation outcomes.

| Office product | Architecture | Functional smoke | Clean install / upgrade |
| --- | --- | --- | --- |
| Microsoft 365 | x86 | Pending | Pending |
| Microsoft 365 | x64 | Pending | Pending |
| Word 2021 | x86 | Pending | Pending |
| Word 2021 | x64 | Pending | Pending |
| Word 2024 | x86 | Pending | Pending |
| Word 2024 | x64 | Local smoke passed; full matrix pending | Pending |

For every product/architecture, exercise mixed-DPI monitors in both directions, several zoom levels, scrolling, multiple documents, focus changes, stale results, Tab/Esc/Down, alternative shortcuts, exact visible-text insertion, one-step Undo, native-prediction coexistence, import cancellation, corrupt/private collection recovery, and source/collection deletion without deleting originals. Check installer prerequisite and third-party-license distribution requirements before shipping.
