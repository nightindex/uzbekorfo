# Morphology evaluation

Baseline date: 2026-09-14  
Rule version: `uzbekorfo-suffixes-v2`  
Corpus version: `2026.09.14.1`

## Accuracy baseline

The engineering seed corpus contains 18 manually inspected forms: 8 correct forms and 10 incorrect forms.

| Metric | Count | Rate |
|---|---:|---:|
| False positives (incorrect form accepted) | 0 / 10 | 0.00% |
| False negatives (correct form rejected) | 0 / 8 | 0.00% |

This is a regression baseline, not a claim of real-world Uzbek accuracy. The corpus is deliberately small and its independent Uzbek-linguist review is still pending. Production decisions must use a larger, genre-balanced corpus after that review.

## Word integration and performance baseline

The isolated Word smoke harness creates a hidden Word instance and checks a 50,000-token document. On the development machine used for this baseline, the final spelling scan completed in 70 ms with zero errors. The release ceiling is 30 seconds to minimize hardware-related flakes.

The workload intentionally repeats a known word, so this figure measures large-document extraction, tokenization, deduplication, and Word integration. It does not represent the cost of generating suggestions for 50,000 distinct errors.

## Reproduce

```powershell
.\eng\test-unit.ps1
.\eng\test-word.ps1
```

The unit test fails if either corpus error count becomes non-zero. The Word test fails if correctness, metadata loading, dual-script morphology, orthographic allomorph checks, or the performance ceiling regresses.
# Scope clarification (2026-09-15)

The historical parser measurements below use a four-root fixture, not the production
dictionary. Their positive class is an accepted form. The new full-engine evaluator
uses an error flagged as positive: false alarms are valid words rejected; missed errors
are invalid words accepted. See [production validation](production-validation.md).
Neither seed corpus is independently reviewed or representative.
