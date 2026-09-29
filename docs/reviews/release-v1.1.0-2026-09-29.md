# Uzbek Orfo v1.1.0 release record

The project owner requested publication without waiting for the remaining external release checks. This record does not mark those checks as passed or claim independent linguistic certification.

## Included changes

- MatnAI document learning reports malformed files individually, keeps rename and source removal changes isolated until save succeeds, restores damaged collections from valid backups, and improves import feedback and cancellation.
- Ghost suggestion positioning uses the Word caret top when the caret is taller than the text cell.
- The About page lets wrapped text determine its height at different display scales.
- The experimental next-100 LexUZ candidate sources and models are not included in the installed application.

## Checks completed on this candidate

- Repository preflight passed, with the linguistic approval warning still present.
- Signed Release build and offline ClickOnce publish completed with the existing local certificate.
- 115 unit tests passed; the UI compatibility harness passed through 384 DPI.
- Isolated Word smoke checks passed on Word 16.0 build 16.0.17932. The harness was x64; Office bitness was not established by that check.
- The generated `setup.exe` has a signer, but Windows does not trust its self-signed certificate chain on this machine. Users on other machines should expect a publisher trust warning.

## Checks not completed

- `docs/reviews/linguist-approval.json` remains pending. There is no independent Uzbek reviewer sign-off for the dictionary, rules, or prediction contexts.
- All six clean installation and upgrade rows in `docs/reviews/office-installation-matrix.json` remain pending. The Word smoke check is not an installation or upgrade test.
- Signing trust, timestamping, and upgrade identity have not been verified on target machines.

The release is distributed with these known validation limits. A later update should complete the pending reviews and installation matrix before claiming production readiness.
