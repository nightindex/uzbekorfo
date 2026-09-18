# Selective-suggestion pilot — 2026-09-17

Status: implemented for evaluation, not production certification. No independent
lawyer pilot or legal-language review has been completed. No neural model is used.

## What changed

After a completed word and space, a candidate's first word must have at least two
observations, at least 40% of next-word support, and a lead of at least 10 percentage
points over the runner-up in the matched context. The gate applies before personal
acceptance boosts and to multiword suggestions too. Context backoff remains enabled.
Two-letter prefix completion and existing morphology validation are unchanged.
These are pilot thresholds, not independently approved release thresholds.

Compare the previous policy and the selective policy without replacing packaged
models or the human-review templates:

```powershell
./eng/build-prediction-corpus.ps1 -Script cyrillic -PilotEvaluation
./eng/build-prediction-corpus.ps1 -Script latin -PilotEvaluation
```

Both commands require the local source snapshots. Cyrillic inputs now use
`corpus/raw_cyrillic/`, with legacy `corpus/raw/` as a fallback. Latin inputs use
`corpus/raw-latin/` and its manifest. No documents are removed or downloaded by
this evaluation.

## Controlled replay results

Each script has 900 held-out after-space cases. Values below are previous → selective.

| Script | Exact top-1 among offered suggestions | Coverage | Simulated remaining keystrokes saved |
| --- | ---: | ---: | ---: |
| Cyrillic | 21.5% → 30.3% | 83.3% → 52.8% | 12.4% → 11.4% |
| Latin | 20.4% → 28.5% | 78.9% → 43.7% | 11.1% → 9.1% |

This is a selectivity tradeoff: fewer interruptions and higher exact-match rate
among suggestions offered, but lower simulated keystroke savings. Overall top-1
across all 900 cases also decreases (Cyrillic 17.9% → 16.0%; Latin 16.1% → 12.4%).
It is not an across-the-board accuracy improvement or evidence of actual time saved.

The 600 two-letter-prefix cases per script are unchanged: Cyrillic top-1 41.5%,
top-3 53.8%; Latin top-1 35.7%, top-3 49.0%. See the full
[Cyrillic report](pilot-report.json) and [Latin report](latin/pilot-report.json),
including separate validation results.

Replay uses training-only models and the same normalized matching for both
policies. Related versions and duplicates are grouped before splitting; training
sentences are excluded from replay. Grouping is heuristic and requires independent
leakage review. Exact text match does not establish legal correctness or usefulness.
The older `corpus-report.json` files describe the earlier policy.

## Performance scope

The memory optimization shares immutable provenance by source/support, reuses
equal strings within each index snapshot (never through global string interning),
and compacts provenance lists to arrays. No thresholds, model files, training
counts or ranking rules changed. All recorded validation/test replay metrics in
both scripts were identical before and after. See the
[before/after measurements](memory-comparison.json).

| Both built-in models, managed retained memory | Before | After |
| --- | ---: | ---: |
| Phrase indexes, source models released | 79.8 MiB | 38.0 MiB |
| Phrase indexes plus full source models | 145.1 MiB | 103.4 MiB |

Settings now retains lightweight built-in metadata and precomputed counts rather
than full built-in phrase payloads. Source paths and word/phrase counts remain
visible; editable private collections keep their full data for refresh/removal.
Loading/import still has transient allocations. The table measures engine/model
states, not settings control allocations or peak load memory.

The 100 MB additional-loaded-memory target is still not proven across workflows.
Dictionary data, private collections, native allocations and peak import/index-build
memory are excluded. The source-model-plus-index state still exceeds the target.

The four-query warm smoke workload measured p95 below 1 ms on .NET 8, on a
16-logical-processor, 16 GB development machine. This is not a broad workload,
the four-core reference PC, or Word's .NET Framework runtime. Real Word context
capture in 100,000 words measured p95 15.32 ms in the post-optimization smoke run; this is not
typing-to-ghost latency.

## Opt-in local pilot counters

Enable local aggregate statistics in settings only with the participant's consent.
Ordinary typing and document text are not recorded. Accepted-suggestion learning
is a separate opt-in and is not required for metrics.

Close Word normally, then inspect the latest session:

```powershell
./eng/report-prediction-pilot.ps1
# Or report a saved session snapshot:
./eng/report-prediction-pilot.ps1 -MetricsPath 'C:\Pilot\session-metrics.tsv'
```

The script reads `%LocalAppData%/UzbekOrfo/MatnAI/metrics-v1.tsv` and prints JSON;
it does not upload anything. Preserve an authorized copy of each session snapshot
before the next session overwrites it. Output contains counts and ratios only,
with null ratios when no denominator exists. It rejects malformed, negative,
duplicate or unknown counters. Do not combine overlapping snapshots as new sessions.

Counters distinguish query attempts, shown suggestions, acceptances, explicit
dismissals, edits/caret moves, no-candidate results, presentation failures,
alternatives opened, insertion rejection and immediate keyboard Undo. Automatic
cleanup after acceptance/settings changes does not count as explicit dismissal.
Counts are event totals, not a one-to-one conversion funnel: some requests become
stale, and a presentation may be replaced or hidden without explicit dismissal.

Limitations:

- Edits and caret moves are combined; they are not necessarily rejections.
- Immediate Undo covers detected Ctrl+Z within five seconds, not all Undo paths.
- Latency buckets begin at detected context change, excluding polling delay.
- Counters alone cannot measure task time, legal correctness or cognitive load.
- Older snapshots may contain dismissal counts from automatic cleanup; do not mix
  those with sessions collected by this version. Record the tested build separately.

## Pilot protocol and release gates

1. Recruit 5–10 consenting Uzbek legal writers. This is a proposed pilot, not a
   claim that participants already exist. Obtain permission before using their
   drafts/contracts; public legislation alone is not representative of their work.
2. Prepare matched drafting/editing tasks in both scripts. Counterbalance the
   order of suggestions-on and suggestions-off sessions. Keep practice tasks
   separate; record build, settings, Office version, zoom and monitor scaling.
3. Measure completion time including proofreading, remaining errors, acceptance,
   explicit dismissal and immediate-Undo counts. Ask about distraction and useful
   versus misleading continuations. Never treat a prediction as legal advice.
4. Independently review at least 200 held-out contexts, with reviewer identity,
   permission/provenance, plausible alternatives and decisions. Keep signed review
   records separate from generated templates. Tune only on validation data, freeze
   thresholds, then evaluate on an untouched test set.
5. Test ghost placement/blinking on physical mixed-DPI monitors, fonts, zoom,
   scrolling, focus changes and multiple documents. Run
   `./eng/test-word.ps1 -Configuration Debug -VisualGhost` on an unlocked desktop
   with its temporary Word window visible and unobstructed. The harness makes one
   foreground-activation attempt (with a briefly attached input queue if needed), then waits up to 30 seconds for you to select
   its temporary Word window. Do not switch applications until it finishes. The
   standalone visual branch uses per-monitor-aware physical coordinates and native
   layered-window capture; it does not change Office's process DPI policy. A background-window
   screenshot is not valid alignment evidence. Do not use an unrelated window's
   matching pixels as a pass.
6. Test Microsoft 365, Word 2021 and Word 2024, both Office architectures, including
   clean installation, upgrade, single-step Undo and native-prediction coexistence.
   Do not silently change Word settings. Pause the pilot for unintended insertions,
   damaged content, focus theft or reproducible editing disruption.
7. Verify warm query p95 ≤50 ms, actual typing-to-ghost p95 ≤150 ms, and additional
   loaded prediction memory ≤100 MB on the four-core/8 GB/SSD reference PC, including
   import and a 100,000-word document. Profile total and peak memory, not just indexes.

## Current verification and known blockers

- 102 unit tests passed, including the ambiguous-after-space gate, prefix fallback,
  shared-provenance support counts and index independence from mutable source data.
- Collection-settings UI tests passed in light and dark modes, including metadata-only
  built-ins, accurate displayed counts/source paths, intact private imports, refresh,
  document selection and 96/144/192/96 DPI transitions. These automated layout checks
  do not validate physical mixed-monitor ghost placement.
- Both 57-file corpora completed the comparison run; packaged models and human
  review templates were not rewritten by the pilot command.
- Real Word nonvisual smoke passed, including exact phrase insertion, one-step Undo,
  stale-context rejection and per-document collection isolation.
- The missing-preview cause was reproduced: the non-activating owned overlay could
  remain behind Word despite having correct geometry and visible-window state.
  Raising the overlay without activation resolved it. The production renderer now
  adjusts z-order after publishing a surface, only while its Word owner is foreground.
  It neither makes the overlay globally topmost nor changes keyboard focus; cached
  idle presentations do not repeatedly raise or repaint it.
- The production fix passed all 24 real Word visual cases on installed Word 16.0,
  build 16.0.17932, at reported window DPI 96: Arial/Calibri/Times New Roman,
  Latin/Cyrillic, 100%/150% zoom, native-caret/fallback anchors. Preview and accepted
  ink top/bottom matched within one pixel; preview ink was required to be nonempty.
  The diagnostic test-only z-order workaround was removed before this passing run.
  Screenshots are in the ignored `tests/WordSmoke/bin/Debug/` output directory.
  The harness still rejects obstructed captures. This does not certify other DPI
  settings, physical mixed-monitor transitions, long-session blinking or all Office
  versions. Existing documents were not modified; temporary test documents closed
  without saving.
- The full UI script stops at its existing ribbon-label assertion because the
  current designer uses `Тезкор Матн` while the assertion requires `MatnAI`. This
  concurrent branding edit was preserved; the full UI run is not reported as passing.
- Independent linguistic review, real-user time savings, the reference-PC targets
  and the complete Office installation matrix remain open.
