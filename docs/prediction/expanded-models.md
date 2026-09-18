# Expanded legal prediction candidates — 2026-09-18

These are statistical word/phrase models, not neural fine-tunes. After the
controlled comparison, both expanded artifacts were promoted to the add-in's
embedded resources on 2026-09-18. Each combines the original corpus with admitted
court documents; Latin and Cyrillic remain separate resources loaded together.
Reproducible candidate outputs also remain under `corpus/candidates/`.
Verified recovery copies of the previous models are in
`.local-archive/model-promotion-20260918-095937/`.

| Measure | Cyrillic | Latin |
| --- | ---: | ---: |
| Original sources | 57 | 57 |
| Additional admitted court sources | 98 | 97 |
| Admitted sources total | 155 | 154 |
| Extracted unique-sentence word count, summed per source | 1,167,435 | 1,169,255 |
| Retained sequence records | 163,358 | 163,054 |
| Held-out cases | 4,500 | 4,169 |
| Original-only training top-1 | 17.93% | 17.61% |
| Expanded training top-1 | 22.71% | 21.52% |
| Original-only training top-3 | 22.76% | 22.07% |
| Expanded training top-3 | 27.64% | 26.34% |
| Original-only simulated keystrokes saved | 9.71% | 10.30% |
| Expanded simulated keystrokes saved | 12.64% | 12.61% |
| Warm query p95 on development machine | 2.12 ms | 1.75 ms |

## Regression slices and promotion decision

| Held-out subset | Cyrillic original → expanded top-1 | Latin original → expanded top-1 |
| --- | ---: | ---: |
| Original documents | 27.89% → 27.44% (900 cases) | 17.72% → 19.94% (1,800 cases) |
| Court documents | 15.44% → 21.53% (3,600 cases) | 17.52% → 22.71% (2,369 cases) |
| After-space prediction | 10.41% → 14.70% | 10.44% → 14.46% |
| Two-letter completion | 35.48% → 41.41% | 32.83% → 36.51% |

Promotion accepts a small Cyrillic original-document regression: four fewer
top-1 matches (251 → 247), and top-3 falls from 35.22% to 34.11%. Simulated
keystroke savings on that subset remain essentially unchanged, 15.06% → 15.07%.
Overall accuracy, coverage and keystroke savings improve in both scripts.
No thresholds were tuned on these test results. This is an engineering trade-off,
not proof of statistically significant improvement or independent quality approval.

Both training variants were evaluated on the same cases for each script, after
grouping duplicates and related versions. The original-only variant is retrained
on original documents belonging to that expanded split's training groups; this
is not a replay against the shipped artifact, which could contain held-out text.
Historical reports have different splits and are not directly comparable.
Cases mix after-space prediction and two-letter completion, with dictionary
fallback. These are not pure next-word-only accuracy figures. Timing is a
development-machine query measurement, not typing-to-ghost latency in Word or
a reference-PC certification. The two script jobs ran concurrently.

The Latin export `sud-faoliyatiga-tegishli/lexuz-3515278.doc` was excluded by the
conservative Latin-only check. It remains on disk for review. Cyrillic checks
require distinctive Uzbek letters but do not prove every passage is Uzbek.
No linguistic approval is claimed, and the download snapshot is not a guarantee
of complete or current coverage of the court website.

Detailed provenance, exclusions, split assignments, metrics and pending review
contexts:

- [Cyrillic report](expanded/corpus-report.json)
- [Latin report](latin/expanded/corpus-report.json)
- [Build instructions](../../corpus/README.md#expanded-candidate-models)

Validation: corpus tool builds successfully; existing unit suite passes 102/102.
Candidate models pass `CollectionStore.Validate` before atomic publication.
Independent review of the held-out contexts and Word pilot testing remain
required before a production-quality claim. No source documents were deleted.
