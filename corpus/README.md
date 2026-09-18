# Model training inputs

Keep the downloaded legal documents if you intend to rebuild, evaluate, or improve
the models. They are source snapshots; downloading a changed law later may not
reproduce the same model.

| Location | Purpose | In Git? |
| --- | --- | --- |
| `raw_cyrillic/` | 57 Cyrillic legal DOC exports (legacy `raw/` also supported) | No |
| `raw-latin/` | 57 Latin exports and their URL manifest | No |
| `raw_cyrillic/sud-court/` | Additional Supreme Court court-materials exports and manifest | No |
| `raw-latin/sud-court/` | Additional Uzbek Latin court-materials exports and manifest | No |
| `../src/UzbekOrfoAddIn/Data/legal_prediction.collection.gz` | Built-in Cyrillic model | Yes, when staged/committed |
| `../src/UzbekOrfoAddIn/Data/legal_prediction_latin.collection.gz` | Built-in Latin model | Yes, when staged/committed |
| `../docs/prediction/` | Source provenance, evaluation, and review contexts | Yes, when staged/committed |

Normal add-in builds use the generated models and do not need raw documents.
Corpus builds require the raw documents at these paths:

```powershell
./eng/build-prediction-corpus.ps1 -Script cyrillic
./eng/build-prediction-corpus.ps1 -Script latin
```

To reclaim disk space later, first back up both raw directories (including the
Latin manifest) outside the repository and verify the backup. Restore them to
these paths before rebuilding. Deleting the only copy would lose the reproducible
training snapshot. Neither corpus is included in the Word add-in package.

The additional court-materials download is documented in
[`docs/legal-sud-court-corpus.md`](../docs/legal-sud-court-corpus.md). It is a
separate source snapshot. The 2026-09-18 combined built-in models include 98
additional Cyrillic sources and 97 Latin sources; one Latin file remains excluded.
See [quality comparison and promotion](../docs/prediction/expanded-models.md).

## Expanded candidate models

Build candidates from the original documents plus the downloaded court materials:

```powershell
./eng/build-prediction-corpus.ps1 -Script cyrillic -IncludeCourt
./eng/build-prediction-corpus.ps1 -Script latin -IncludeCourt
```

These commands write models to `corpus/candidates/` (ignored by Git), not the
add-in's embedded resources. Reports and 200 pending review contexts per script
go to `docs/prediction/expanded/` and `docs/prediction/latin/expanded/`.
These build commands alone do not change embedded model files or original review
records. Promotion is a separate, explicit step after comparing the reports.

The tool checks the nested court manifests and reports excluded court files.
Latin sources containing Cyrillic text are conservatively excluded. Cyrillic
court sources must contain distinctive Uzbek letters; this is a heuristic, not
proof that every passage is Uzbek. Source documents are never modified or deleted.

`originalCorpusSameHoldout` compares original-only training with expanded training
using identical held-out cases and duplicate/version groups. Historical reports
use different splits and must not be treated as an equivalent comparison.
Inspect accuracy, coverage and keystroke savings before promoting a candidate.
This rebuild learns phrase statistics; it is not neural-network fine-tuning.
