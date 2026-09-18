# Production readiness — 2026-09-18

Decision: **not ready for a public production release**. The combined prediction
models are engineering/pilot candidates. No production tag, upload or installer
release is authorized by the user's conditional instruction until readiness is
established. A passing build is not a production approval.

## Confirmed evidence

- Repository preflight passes; generated dictionary data is synchronized.
- Combined-model integration tests verify both scripts, provenance, model
  validation, original-source retention and the 155/154 source counts.
- This audit passed all 102 unit tests in Release configuration, the Release
  add-in build, and all eight synthetic approval-gate checks. These do not test
  clean installation or publisher trust on another machine.
- Grouped held-out replay improves overall prediction accuracy in both scripts;
  a small original-Cyrillic regression remains explicitly documented in
  [the comparison](../prediction/expanded-models.md).
- `validate-linguist-review.ps1 -RequireApproved` rejects the current pending
  approval. This is an expected release blocker, not a test to bypass.

## Unresolved release requirements

1. Independent Uzbek linguistic review, bound to the exact dictionary/rule/corpus
   hashes, plus review of the prediction held-out contexts. No reviewer has signed.
2. Clean installation, upgrade and user-data retention tests on Microsoft 365,
   Word 2021 and Word 2024, each x86 and x64. All six matrix rows are pending.
3. Lawyer pilot evidence for acceptance, dismissal and immediate Undo; mixed-DPI,
   zoom, scrolling, focus and native-prediction coexistence checks.
4. Full end-to-end latency and additional-memory measurements on the specified
   reference machine, including import and a 100,000-word document. Engine-only
   development-machine measurements do not satisfy these requirements.
5. Verify production signing trust, expiry and timestamping; build and validate
   the exact installable package, record its hash and preserve upgrade identity.
   Successful development signing does not prove publisher trust on user machines.
6. Review the dirty worktree before a release commit. It includes substantial
   pending implementation changes and unrelated existing deletions; do not stage
   everything blindly. Publish from a reviewed, clean release branch.

## Repository organization

Keep source code in `src/`, automated tests/fixtures in `tests/`, developer commands
in `eng/`, the corpus builder in `tools/`, and provenance/reviews in `docs/`.
Downloaded source snapshots and candidate models under `corpus/` are ignored;
only bundled model artifacts belong in application resources. Build outputs,
local recovery copies and signing material remain ignored.

Do not remove corpus snapshots, model rollback copies, third-party notices,
Hunspell source provenance, or pending approval records merely to reduce file
count. No additional file deletions were justified during this audit.

## Next decision

Obtain the external review and target-machine evidence, then rerun release
verification. A separately labeled prerelease for testers is a different product
decision and must not be presented as a production-ready installer. No approval
records have been fabricated or marked passed by this audit.
