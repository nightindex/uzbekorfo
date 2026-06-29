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
1. Open `UzbekOrfoAddIn.csproj` in Visual Studio.
2. Build the project (Debug or Release).
3. Start debugging to launch Word with the add-in.

## Project Structure
- `Core/`        Interfaces and core contracts
- `Data/`        Embedded dictionaries, rules, and datasets
- `Forms/`       WinForms dialogs
- `Helpers/`     Utilities and shared helpers
- `Models/`      DTOs and domain models
- `Services/`    Business logic and workflows
- `UI/`          Theming and custom controls
- `Properties/`  Assembly/resources settings

## Notes
- `Data/` files are embedded as resources and copied to output for debugging.
- Signing keys (`*.pfx`) and local `Directory.Build.props` are excluded via the repo `.gitignore`. Copy `Directory.Build.props.example` locally and set signing values when creating release builds.

## Publishing
The project supports ClickOnce publish output (see `publish/` after build). This folder is intentionally ignored for GitHub.
