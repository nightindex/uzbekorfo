# Documentation

## Development and product behavior

- [Architecture](architecture.md)
- [MatnAI prediction, collections, and shortcuts](prediction/README.md)
- [Completion and Word integration details](matnai-mvp.md)
- [Display and DPI compatibility](display-compatibility.md)
- [Document editing safety](review-fixes.md)
- [Release process](release.md)

## Dictionary and source data

- [Dictionary schema](../src/UzbekOrfoAddIn/Data/SCHEMA.md)
- [Imported dictionary/suffix data and recovery](hunspell-import.md)
- Legal corpus provenance: [codes](legal-corpus.md), [laws](legal-laws-corpus.md),
  [government documents](legal-government-corpus.md), [Latin model](legal-latin-corpus.md)
- [Training inputs and retention](../corpus/README.md)

Hunspell is not a runtime dependency. Its imported root flags still constrain
suffix compatibility, and `reference/hunspell_sources.json.gz` retains the six
original inputs for audit and recovery. Import/restore scripts are maintenance
tools, not components loaded in Word.

## Quality evidence

- [Production validation](production-validation.md)
- [Linguist review](linguist-review.md) and [morphology evaluation](morphology-evaluation.md)
- [Prediction evaluation and outstanding checks](prediction/validation.md)
- Current bundled models: [comparison and regression slices](prediction/expanded-models.md),
  [Cyrillic report](prediction/expanded/corpus-report.json), [Latin report](prediction/latin/expanded/corpus-report.json)
- [Production readiness audit, 2026-09-18](reviews/production-readiness-2026-09-18.md)
- `reviews/`: approval records, Office installation matrix, and confirmed dictionary corrections.

These records remain necessary even when their status is pending. Removing them
does not complete the corresponding release checks. The retired completion plan
and one-time bundle migration script are superseded by the current documents above.

## Local files

Generated `bin/`, `obj/`, `.vs/`, and `TestResults/` directories are ignored by Git.
`TestResults/` can contain dictionary recovery copies and human review work:
inspect it before cleaning. `.local-archive/` holds recoverable cleanup copies and
is also ignored. Signing keys and local build properties are not cleanup targets.
