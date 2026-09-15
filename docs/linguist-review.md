# Uzbek morphology linguist review

Production release is intentionally blocked until an independent Uzbek linguist reviews the versioned morphology rules and the regression corpus.

Use the [production validation procedure](production-validation.md) to export review
worksheets, measure the full spelling engine and record hash-bound approval. The
[2026-09-15 engineering audit](reviews/engineering-review-2026-09-15.md) is not a
linguist sign-off. Approval now requires `docs/reviews/linguist-approval.json` in
addition to corpus review metadata; changing file content invalidates its hashes.

## Review package

- `src/UzbekOrfoAddIn/Data/uzbek_suffixes.json` contains suffix families, ordering stages, compatibility constraints, allomorph groups, and root mutations.
- `tests/Corpora/uzbek_morphology_corpus.json` contains accepted and rejected examples and records the reviewed rule version.
- `docs/morphology-evaluation.md` records measured false-positive and false-negative rates.

## Required review

The linguist should verify:

1. Every suffix surface form in Latin and Cyrillic.
2. Family and part-of-speech compatibility.
3. Ordering stages and maximum occurrence groups.
4. Dative `-ga/-ka/-qa`, invariant written locative `-da`, and invariant written ablative `-dan`.
5. Vowel-final possessive forms and consonant-final `k→g`, `q→g‘` mutation, including lexical exceptions.
6. Verb negation, tense/gerund/imperative cores, and agreement combinations.
7. Every corpus label, plus additional counterexamples from edited Uzbek prose.

The current engineering correction for `-da/-dan` is supported by an [official government assessment table](https://api-portal.gov.uz/uploads/9/2026/04/16/96732824-a6f3-f6a1-d46e-2101eecc1f48_media_.pdf). The broader baseline follows the published [O‘zbek tilining asosiy imlo qoidalari](https://imlo.uz/collections/ozbek-tilining-asosiy-imlo-qoidalari). These references support engineering work but do not replace named human approval.

## Approval procedure

After review, set these fields in the corpus:

```json
"linguistStatus": "approved",
"linguistName": "Reviewer’s full name",
"reviewedAtUtc": "YYYY-MM-DDTHH:MM:SSZ"
```

Then run:

```powershell
.\eng\validate-linguist-review.ps1 -RequireApproved
.\eng\test-unit.ps1
.\eng\test-word.ps1
```

Any subsequent rule-version change resets approval to `pending` and requires another review.
