# MatnAI accuracy candidate

The alternative ranking policy is implemented for evaluation and is not enabled
in Word. The existing policy remains active because the candidate failed the
agreed automated gate on the held-out formal Uzbek corpus. Independent review
of the 200 contexts for each script is also pending.

Run the separate Cyrillic and Latin comparisons without replacing model files
or generated review templates:

```powershell
./eng/build-prediction-corpus.ps1 -Script cyrillic -AccuracyEvaluation
./eng/build-prediction-corpus.ps1 -Script latin -AccuracyEvaluation
```

Reports and reviewer packets go to `TestResults/MatnAiAccuracy-implementation/`.
The packet includes the source, context, expected word, both policies' top three
suggestions, and fields for reviewer judgment. Keep completed judgments in a
separate signed file; rerunning evaluation replaces the generated packet.

## Initial held-out comparison

| Script | Exact top-1 among offered, current → candidate | Coverage, current → candidate | Simulated keystrokes saved, current → candidate |
| --- | ---: | ---: | ---: |
| Cyrillic | 36.6% → 35.9% | 71.7% → 65.1% | 15.6% → 15.6% |
| Latin | 32.8% → 30.9% | 66.2% → 64.0% | 13.0% → 13.7% |

The candidate improved after-space precision but reduced prefix precision.
The reported query p95 covers the local engine, dictionary lookup, and ordering;
it does not measure Word UI latency or morphology validation. No reviewer has
judged whether either suggestion is useful or misleading in context.

The release gate requires at least a three-point exact top-1 improvement in
both scripts, no more than a five-point coverage loss, no decline in simulated
keystroke savings, no measured latency regression, and an independent Uzbek
review of the held-out contexts. The current candidate does not qualify.
