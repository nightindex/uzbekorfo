# Release process

A production release is a signed ClickOnce publication built from a clean `release/*` branch. The release command is deliberately strict: it blocks unsigned publishing, dirty source trees, non-release branches, unsupported update channels, failed automated checks, and failed Word smoke tests.

## One-time signing setup

1. Obtain a trusted code-signing certificate, or obtain an approved open-source signing service.
2. Keep its private key outside Git. Copy `Directory.Build.props.example` to the ignored root `Directory.Build.props` and set `ManifestKeyFile` and `ManifestCertificateThumbprint`. The certificate file must be accessible to the account running the build.
3. For CI, set the protected `UZBEKORFO_MANIFEST_KEY_FILE` and `UZBEKORFO_MANIFEST_CERTIFICATE_THUMBPRINT` environment variables after your signing provider has made the certificate available to the build agent. Never add a certificate, password, or thumbprint to a committed project file.

The `Publish` target refuses an unsigned Release publish. A normal local Debug build may remain unsigned for development only.

## Update channels

### In-app GitHub check

`Маълумот → Дастур ҳақида → Янгиланишни текшириш` explicitly requests
`https://api.github.com/repos/nightindex/uzbekorfo/releases/latest`.
No startup polling, document content, credentials, or automatic installer
downloads are involved. GitHub receives ordinary network request metadata.
The request times out after 15 seconds. Failure leaves the browser link available.

Publish stable releases with tags `vMAJOR.MINOR.PATCH` or
`vMAJOR.MINOR.PATCH.REVISION`, matching the ClickOnce `ApplicationVersion`.
The installed ClickOnce version is preferred; development builds use the assembly
version. Drafts, prereleases, and unsupported tag formats are not offered.
Release notes are shown as plain text, never executed or rendered as HTML.
The latest-release API contract is documented in
[GitHub's release API documentation](https://docs.github.com/en/rest/releases/releases#get-the-latest-release).

`Юклаб олиш саҳифаси` opens the fixed official repository's release page.
Users download the complete signed installer ZIP (not GitHub's source-code ZIP),
extract it, save documents and close Word before installing. The same package
can be transferred to an offline PC; .NET Framework and VSTO prerequisites must
already be installed or supplied separately. Preserve ClickOnce application and
signing identity between versions and test settings/collection retention during
upgrade. This feature does not build or publish a release by itself.

`Offline` is the default channel. Each release produces a signed package that users install manually; updates are distributed by publishing the next package.

`Web` enables ClickOnce foreground updates. It requires a stable HTTPS URL that you control, such as `https://downloads.example.com/uzbekorfo/`. Keep every published deployment manifest and application-files directory reachable at that URL for existing users. Do not enable this channel until the download hosting and retention policy are ready.

## Local certificate with GitHub Releases

For a no-cost open-source release, you may sign with a local/self-signed certificate and use the `Offline` channel. This proves that files in one package came from the same local key, but it **does not make the publisher trusted**: Windows will show an Unknown Publisher or trust warning for users who do not install and trust your certificate themselves.

Use this only when you explicitly accept that trade-off:

```powershell
.\eng\release.ps1 -Channel Offline -AllowUntrustedCertificate
```

The switch is deliberately required for a temporary or test-named certificate. Never commit or upload the `.pfx`, its password, or any private key. Upload only a ZIP of the complete generated offline publish directory to a GitHub Release; users must extract it before starting the installer. Each new GitHub Release is a manual update for users.

## Release command

After all changes are reviewed and committed on a branch such as `release/v1.1.0`, run:

```powershell
.\eng\release.ps1 -Channel Offline
```

For a hosted update channel:

```powershell
.\eng\release.ps1 -Channel Web -UpdateUrl 'https://downloads.example.com/uzbekorfo/' -PublishDirectory 'C:\release\uzbekorfo'
```

The command runs repository preflight, UI compatibility tests, the complete unit-test suite, a signed Release publish, and `eng\test-word.ps1`. The Word step creates an isolated, hidden Word instance and closes all of its temporary documents without saving.

## Manual release record

Complete the six-target clean installation and upgrade matrix in
[production validation](production-validation.md). A Word smoke pass is not proof of
installation or upgrade compatibility. Independent linguistic approval must match the
exact dictionary, suffix rules and corpus hashes.

Before shipping, record the version, commit SHA, certificate publisher, channel URL (if any), Windows version, and pass/fail result for each supported Word version. On physical target devices, verify the DPI/accessibility checks in [display compatibility](display-compatibility.md), including keyboard navigation and Narrator labels at the intended display scales.

The release script cannot replace these physical-device checks or obtain a certificate; they are external release approvals.
