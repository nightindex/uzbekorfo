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
2. Build the project (Debug or Release).
3. Start debugging to launch Word with the add-in.

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
The project supports ClickOnce publish output (see `publish/` after build). This folder is intentionally ignored for GitHub.

## Validation

Run the repository checks from the root:

```powershell
.\eng\preflight.ps1
dotnet test .\tests\UzbekOrfoAddIn.UnitTests\UzbekOrfoAddIn.UnitTests.csproj
```

The test project covers pure helper logic and does not require Word. VSTO build validation requires a Windows machine with the Office development workload installed.
