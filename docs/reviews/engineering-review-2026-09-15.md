# Engineering review: 2026-09-15

Not independent linguistic approval. This review did not change dictionary words or
suffix behavior and did not install, upgrade or uninstall Office/the add-in.

## Dictionary and suffix screening

Source dictionary SHA-256: `0DDAA9ECDC7296BE37FE211EA50203077BAF8AAA9F685B60EBEBCF192DDCEBE0`.
Suffix JSON SHA-256: `817FC1B59DDCDD2B920B3D0E77103B20AD3E25132906BC5B9F5F8C23D9A1F0F9`.

- 216,341 entries screened; 187,107 carry source flags.
- 216,317 entries lack explicit lemma and POS fields. Inferred/fallback metadata is
  not independently reviewed annotation. Prioritize POS-sensitive suffix compatibility.
- 3,325 entries flagged: 3,012 whitespace/invisible-character only, 295 normalized
  duplicate only, one with both findings, and 17 mixed Latin/Cyrillic entries.
  These are not 3,325 confirmed spelling errors. Sample whitespace findings include
  dictionary phrases such as `abri bahor`; review phrase lookup/tokenization separately
  instead of deleting legitimate multiword entries.
- 2,384 suffix rules exported for review, including 2,333 imported complete forms.
  Source flags and structural compatibility do not establish linguistic correctness.

The four-root parser fixture exercises no imported source flags. Its missing-root
negative tests a fixture boundary, not necessarily an invalid Uzbek word. Do not reuse
that label to report full-dictionary accuracy.

## Measurements actually run

| Check | Result | Limitation |
| --- | --- | --- |
| Unit suite | 51 passed, zero failed | Existing compiler nullability/platform warnings remain |
| Approval-gate synthetic tests | 8 passed | Tests stale hashes/metadata, not reviewer authenticity |
| Full-dictionary spelling evaluator | False alarms 0/4; missed errors 0/2 | Six provisional smoke-derived forms, not independent or representative |
| Local Word COM smoke | Passed, including fresh seeding and idempotent suffix data migration | Does not test ClickOnce installation or upgrade |
| 50,000 repeated known-word tokens in Word | 113 ms | Not varied-document or interactive latency evidence |

Local Word reports version 16.0, build 16.0.17932. Click-to-Run configuration reports
16.0.17932.20574, x64, with ProPlus2024Volume among installed product IDs. Shared Office
version 16.0 must not be interpreted as testing all Word editions. All six clean
installation/upgrade matrix rows remain pending.

## Unresolved release dependencies

Obtain an independent Uzbek linguist and permission-cleared representative corpus;
adjudicate the review worksheets and evaluate a held-out set. Run the signed release
candidate and previous-release upgrade on all six agreed Office targets in isolated
environments. See [validation procedure](../production-validation.md).
