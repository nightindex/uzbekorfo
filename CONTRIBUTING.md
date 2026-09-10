# Contributing

## Prerequisites

Use Windows with Visual Studio, the Office/VSTO development workload, .NET Framework 4.7.2 targeting support, and a compatible Microsoft Word installation.

## Before opening a pull request

1. Run `./eng/preflight.ps1` from the repository root.
2. Run `dotnet test ./tests/UzbekOrfoAddIn.UnitTests/UzbekOrfoAddIn.UnitTests.csproj`.
3. Build and smoke-test the add-in in Word when changing VSTO, ribbon, or interop code.
4. Do not commit signing certificates, local `Directory.Build.props`, build output, or user data.

## Code organization

Keep reusable, host-independent logic out of forms, ribbons, and Word interop classes whenever practical. Add tests for that logic under `tests/`. Keep UI display concerns in `Forms/` or `UI/`, and keep file/settings concerns in services designed for that purpose.
