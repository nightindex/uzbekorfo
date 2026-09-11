# Uzbek Orfo Add-in

Uzbek Orfo is a Microsoft Word VSTO add-in for Uzbek spelling/grammar assistance, dictionary management, and script conversion.

## Features
- Spell and grammar checking with error lists
- Suggestions and replace-all workflows
- Personal dictionary management (add/import/export)
- Definitions/explanations lookup
- Latin - Cyrillic transliteration helpers
- Export utilities (errors, reports)

## Requirements
- Windows
- Microsoft Word (Office 2016+ recommended)
- Visual Studio with Office/VSTO tools
- .NET Framework 4.7.2

## Build and Run
1. Open `UzbekOrfoAddIn.slnx` in Visual Studio.
2. Copy `Directory.Build.props.example` to the Git-ignored `Directory.Build.props` and configure your local signing certificate. VSTO requires signed manifests for a runnable add-in, including local debugging.
3. Build the project (Debug or Release).
4. Start debugging to launch Word with the add-in.

The VSTO workload and a compatible local installation of Microsoft Word are required for a full build and debug session.

## Project Structure
- `src/UzbekOrfoAddIn/`  The VSTO host and application source
- `eng/`                Repository validation and engineering scripts
- `tests/`              Automated tests for host-independent logic
- `docs/`               Architecture and contributor documentation
- `.github/workflows/`  Continuous-integration checks

Within `src/UzbekOrfoAddIn/`, `Core/` contains contracts, `Models/` contains domain data, `Services/` contains application logic, `Forms/` and `UI/` contain presentation code, and `Data/` contains bundled dictionaries and rules.

## Notes
- `Data/` files are embedded as resources and copied to output for debugging.
- Signing keys (`*.pfx`) and local `Directory.Build.props` are excluded via the repo `.gitignore`. Copy the root `Directory.Build.props.example` to `Directory.Build.props` locally and set signing values when creating release builds. Do not add certificate files or thumbprints to the project file.

## Publishing
Production publishing is a signed ClickOnce Release build. See [the release process](docs/release.md) for signing setup, offline versus web update channels, the required release branch, and the single release-verification command. Publish output is intentionally ignored by Git.

## Validation

Run the repository checks from the root:

```powershell
.\eng\preflight.ps1
.\eng\test-ui.ps1
.\eng\test-unit.ps1
```

The test project covers helper logic, atomic settings writes, and correction/highlight safety using an in-memory Word double. It does not require Word and does not replace COM integration testing. VSTO build validation requires a Windows machine with the Office development workload installed.

The Windows UI harness requires Visual Studio's C# compiler but not Word. It checks DPI scaling, scrollbar regressions, and small-window layouts. See [display compatibility](docs/display-compatibility.md) for scope and manual checks.

Validate bundled dictionary parity with `./eng/validate-dictionary-sync.ps1`. If the two source files intentionally need to be regenerated, run `./eng/sync-dictionary-data.ps1`; it keeps matching JSON metadata and adds empty metadata fields for spelling-only words.

See [document safety and review fixes](docs/review-fixes.md) for behavior changes and Word smoke tests.
