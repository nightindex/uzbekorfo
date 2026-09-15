# Production validation

Status: engineering preparation only. No independent Uzbek linguist or independently
labelled representative corpus has been supplied. No clean installation/upgrade target
is certified. Passing unit tests or Word COM smoke tests does not change these statuses.

## Linguistic review package

Run `eng/export-linguist-review.ps1`. Local output goes to
`TestResults/linguist-review`: a hash-bound summary, all suffixes in `suffix-review.csv`,
and suspicious dictionary entries in `dictionary-screening.csv`.
Findings are review candidates, not automatic deletions. Duplicate screening normalizes
Unicode, case and apostrophes; it is not the runtime transliterator.

The independent reviewer should adjudicate every imported suffix class, root-flag
compatibility, complete-form boundaries, ordering, allomorphs and lexical exceptions.
Check both scripts, apostrophe variants and hyphenated forms. Include accepted and
rejected examples, their contexts, source references and the reason for each decision.
Review entries lacking explicit POS/lemma separately; source flags are not a substitute
for linguistic metadata. Imported source preservation does not prove linguistic correctness.

## Accuracy measurement

Build the add-in, then run `eng/measure-spelling.ps1 -CorpusPath <reviewed-corpus.json>`.
Use `tests/Corpora/spelling_engineering_seed.json` as the input format, not as a
representative dataset. Every case needs a unique ID, word, boolean expectedCorrect,
category and source. Store context, permission/provenance and reviewer decisions with
the corpus. Output `TestResults/spelling/results.json` contains per-case decisions,
category counts and exact engine/data/corpus hashes. No personal dictionary or unknown-word
collection is used. Exit zero means measurement succeeded, not that accuracy is sufficient.

Positive means **an error flagged**:

- False-positive rate (false alarms): valid forms rejected / all valid forms.
- False-negative rate (missed errors): invalid forms accepted / all invalid forms.

The older parser fixture uses the opposite positive convention (accepted form) and only
four roots. Its artificial missing-root negative must not be used as a real spelling label.
Neither fixture establishes field accuracy. The new evaluator tests word acceptance;
context-sensitive grammar, tokenization, suggestions and Word interaction need separate tests.

Before reporting production accuracy, recruit an independent Uzbek linguist and assemble
permission-cleared edited prose plus independently adjudicated real errors across both
scripts, education, business, news and everyday writing. Keep a document-level held-out
set out of rule development. Report token-weighted and unique-form results separately,
script/domain strata, sample counts and uncertainty; agree release thresholds before
evaluating the held-out set. A tiny zero-error sample is not evidence of a zero error rate.

Record sign-off in `docs/reviews/linguist-approval.json`, including the four SHA-256
hashes computed with `Get-FileHash -Algorithm SHA256`, reviewer identity, independence,
date and evidence reference. Update parser corpus review metadata as required by the
existing gate. For a different evaluation corpus, pass `-EvaluationCorpusPath` to
`eng/validate-linguist-review.ps1` and configure the same path in the release workflow.
The default evaluation corpus remains provisional. Approval is a human attestation;
the script detects stale content but cannot authenticate a reviewer or judge representativeness.

## Clean installation and upgrade matrix

Required scope is Microsoft 365, Word 2021 and Word 2024, each in x86 and x64 Office.
Record results in `docs/reviews/office-installation-matrix.json`; all six rows start pending.
Office architecture is not Windows architecture. Use separate disposable VMs/profiles
with properly licensed Office, a recorded Windows build and an exact Office build/channel.
Do not uninstall or replace a developer's Office installation to simulate this matrix.

For every target:

1. Record candidate version, commit, signed package hash, publisher, OS and Office build,
   Office bitness, prerequisites, tester and UTC date. Snapshot a clean VM/profile.
2. Install the signed distribution as the intended standard user. Record prerequisites,
   publisher/trust prompts, installer logs and result. Do not manually copy DLLs or
   pre-register the development add-in to make the clean-install test pass.
3. Launch Word; verify ribbon, add-in loading, Latin/Cyrillic checks, suggestions,
   corrections, custom dictionary, settings and persistence after closing/reopening Word.
4. Restore a clean snapshot; install the previous shipped version. Add distinctive custom
   words/settings and keep copies/hashes. Close Word, upgrade using the supported channel,
   reopen and verify candidate version, preserved custom data, suffix migration and backup.
   Record the actual previous version; a synthetic legacy JSON fixture is not an upgrade install.
5. Run `eng/test-word.ps1` against the matching built candidate for document-safety smoke
   checks. Also test the installed UI on a varied large document; repeated known-word
   timing alone does not establish interactive performance.
6. Uninstall/reinstall and check expected user-data retention. Attach logs and screenshots
   plus a pass/fail result for every required check. Leave failures and unrun checks visible.

This matrix is a manual release approval, not yet an automated installer test or a
machine-enforced release gate. Do not ship until all six rows have reviewed evidence.
