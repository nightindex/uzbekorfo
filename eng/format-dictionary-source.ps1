<#
.SYNOPSIS
Formats the canonical Uzbek dictionary JSON without changing its data.

.DESCRIPTION
Use after a large dictionary import so uzbek_dictionary.json remains reviewable
and practical to edit. Generated runtime artifacts are not written here; run
sync-dictionary-data.ps1 afterward.
#>
param(
    [string]$DataDirectory = (Join-Path $PSScriptRoot "..\src\UzbekOrfoAddIn\Data")
)

$ErrorActionPreference = "Stop"
$jsonPath = Join-Path $DataDirectory "uzbek_dictionary.json"
if (-not (Test-Path -LiteralPath $jsonPath)) {
    throw "Missing canonical dictionary JSON: $jsonPath"
}

[void][Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
$utf8Strict = New-Object Text.UTF8Encoding($false, $true)
$serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$serializer.MaxJsonLength = [int]::MaxValue

$payload = $serializer.DeserializeObject([IO.File]::ReadAllText($jsonPath, $utf8Strict))
$formatted = ConvertTo-Json -InputObject $payload -Depth 6

$tempPath = "$jsonPath.$([Guid]::NewGuid().ToString('N')).tmp"
try {
    [IO.File]::WriteAllText($tempPath, $formatted + [Environment]::NewLine, (New-Object Text.UTF8Encoding($false)))
    [void]$serializer.DeserializeObject([IO.File]::ReadAllText($tempPath, $utf8Strict))
    Move-Item -LiteralPath $tempPath -Destination $jsonPath -Force
}
finally {
    if (Test-Path -LiteralPath $tempPath) {
        Remove-Item -LiteralPath $tempPath -Force
    }
}

Write-Host "Formatted canonical dictionary source: $jsonPath"
