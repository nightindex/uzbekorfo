# Uzbek Orfo 1.1.0 preview 1 — release preparation

Intended channel: prerelease for testers, not production certification.

Includes document-based Latin/Cyrillic predictions, collection management,
expanded legal models, ghost-text fixes, About-dialog layout fixes, and explicit
GitHub update checking. Prereleases are intentionally excluded from the app's
stable-update checker.

Verified locally: 112 unit tests; UI harness at 96–384 DPI; signed Release build;
Word 16.0 build 16.0.17932 document-safety smoke tests. These are not six-target
clean-install or upgrade certification.

The user reports independent linguistic review passed. Reviewer identity,
review evidence and exact reviewed content hashes have not yet been supplied,
so the repository approval record is not changed by this assertion.

## Proposed assets

- Complete offline application ZIP, including `setup.exe`, the `.vsto` manifest
  and the `Application Files` directory.
- `setup.exe` separately, **not a standalone single-file installer**. It needs
  the extracted application package beside it. Extract the ZIP, save documents,
  close Word, then run setup from that directory.
- Missing .NET Framework/VSTO prerequisites require Internet access; for an
  offline PC, install prerequisites separately first.

The existing local signing certificate is self-signed/untrusted on other
machines. Publishing awaits explicit acceptance of that warning and confirmation
of the EXE format. Do not describe this package as a trusted production installer.
No keys, certificates, downloaded training sources or local recovery files belong
in release assets or the source commit.
