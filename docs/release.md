# Release process

A production release is a signed ClickOnce publication built from a clean `release/*` branch. The release command is deliberately strict: it blocks unsigned publishing, dirty source trees, non-release branches, unsupported update channels, failed automated checks, and failed Word smoke tests.

## One-time signing setup

1. Obtain a trusted code-signing certificate, or obtain an approved open-source signing service.
2. Keep its private key outside Git. Copy `Directory.Build.props.example` to the ignored root `Directory.Build.props` and set `ManifestKeyFile` and `ManifestCertificateThumbprint`. The certificate file must be accessible to the account running the build.
3. For CI, set the protected `UZBEKORFO_MANIFEST_KEY_FILE` and `UZBEKORFO_MANIFEST_CERTIFICATE_THUMBPRINT` environment variables after your signing provider has made the certificate available to the build agent. Never add a certificate, password, or thumbprint to a committed project file.

The `Publish` target refuses an unsigned Release publish. A normal local Debug build may remain unsigned for development only.

## Update channels

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

Before shipping, record the version, commit SHA, certificate publisher, channel URL (if any), Windows version, and pass/fail result for each supported Word version. On physical target devices, verify the DPI/accessibility checks in [display compatibility](display-compatibility.md), including keyboard navigation and Narrator labels at the intended display scales.

The release script cannot replace these physical-device checks or obtain a certificate; they are external release approvals.
