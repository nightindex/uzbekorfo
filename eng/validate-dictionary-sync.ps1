<#
.SYNOPSIS
Verifies that the generated runtime DIC exactly matches the canonical JSON.
#>
param(
    [string]$DataDirectory = (Join-Path $PSScriptRoot "..\src\UzbekOrfoAddIn\Data")
)

$ErrorActionPreference = "Stop"
$generator = Join-Path $PSScriptRoot "sync-dictionary-data.ps1"

if (-not (Test-Path -LiteralPath $generator)) {
    throw "Dictionary generator is missing: $generator"
}

& $generator -DataDirectory $DataDirectory -Check
