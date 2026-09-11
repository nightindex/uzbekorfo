<#
.SYNOPSIS
Generates runtime dictionary artifacts from the canonical JSON dataset.

.DESCRIPTION
uzbek_dictionary.json is the only editable source for bundled main-dictionary
words and their metadata. This script derives the fast spelling list
(uzbek_main.dic) and the compact runtime metadata index
(uzbek_dictionary_metadata.json). Use -Check in CI to verify both artifacts.
#>
param(
    [string]$DataDirectory = (Join-Path $PSScriptRoot "..\src\UzbekOrfoAddIn\Data"),
    [switch]$Check
)

$ErrorActionPreference = "Stop"

$jsonPath = Join-Path $DataDirectory "uzbek_dictionary.json"
$dicPath = Join-Path $DataDirectory "uzbek_main.dic"
$metadataPath = Join-Path $DataDirectory "uzbek_dictionary_metadata.json"
$expectedSchema = "uzbekorfo-dictionary-v1"

if (-not (Test-Path -LiteralPath $jsonPath)) {
    throw "Missing canonical dictionary JSON: $jsonPath"
}

[void][Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
$utf8Strict = New-Object Text.UTF8Encoding($false, $true)
$serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$serializer.MaxJsonLength = [int]::MaxValue

try {
    $payload = $serializer.DeserializeObject([IO.File]::ReadAllText($jsonPath, $utf8Strict))
}
catch {
    throw "Invalid UTF-8 or JSON in canonical dictionary '$jsonPath': $($_.Exception.Message)"
}

if ($null -eq $payload -or -not $payload.ContainsKey("Entries")) {
    throw "Canonical dictionary must contain an Entries array."
}
if (-not $payload.ContainsKey("Schema") -or $payload.Schema -ne $expectedSchema) {
    throw "Canonical dictionary Schema must be '$expectedSchema'."
}

$words = New-Object System.Collections.Generic.List[string]
$metadataEntries = New-Object System.Collections.Generic.List[object]
$seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$entryIndex = 0

foreach ($entry in @($payload.Entries)) {
    $entryIndex++
    if ($null -eq $entry -or -not $entry.ContainsKey("Word")) {
        throw "Dictionary entry #$entryIndex has no Word value."
    }

    $word = ([string]$entry.Word).Trim()
    if ([string]::IsNullOrWhiteSpace($word)) {
        throw "Dictionary entry #$entryIndex has an empty Word value."
    }
    if ($word.StartsWith("#")) {
        throw "Dictionary entry #$entryIndex ('$word') starts with '#', which is reserved for DIC comments."
    }
    if ($word.IndexOfAny([char[]]@([char]13, [char]10)) -ge 0) {
        throw "Dictionary entry #$entryIndex contains a line break."
    }
    if ([regex]::IsMatch($word, "[\p{Cc}\p{Cf}]")) {
        throw "Dictionary entry #$entryIndex ('$word') contains an invisible or control character."
    }
    if (-not $seen.Add($word)) {
        throw "Duplicate dictionary word (case-insensitive): '$word'."
    }

    $words.Add($word)

    $hasMetadata =
        ($entry.ContainsKey("Definition") -and -not [string]::IsNullOrWhiteSpace([string]$entry.Definition)) -or
        ($entry.ContainsKey("SpellingRule") -and -not [string]::IsNullOrWhiteSpace([string]$entry.SpellingRule)) -or
        ($entry.ContainsKey("GrammarNote") -and -not [string]::IsNullOrWhiteSpace([string]$entry.GrammarNote)) -or
        ($entry.ContainsKey("Examples") -and $null -ne $entry.Examples -and @($entry.Examples).Count -gt 0)

    if ($hasMetadata) {
        $metadataEntries.Add([ordered]@{
            Word = $word
            Definition = if ($entry.ContainsKey("Definition")) { $entry.Definition } else { $null }
            SpellingRule = if ($entry.ContainsKey("SpellingRule")) { $entry.SpellingRule } else { $null }
            GrammarNote = if ($entry.ContainsKey("GrammarNote")) { $entry.GrammarNote } else { $null }
            Examples = if ($entry.ContainsKey("Examples")) { @($entry.Examples) } else { $null }
        })
    }
}

if ($words.Count -eq 0) {
    throw "Canonical dictionary contains no words."
}
if ($payload.ContainsKey("WordCount") -and [int]$payload.WordCount -ne $words.Count) {
    throw "Canonical WordCount is $($payload.WordCount), but Entries contains $($words.Count) words."
}
if ($payload.ContainsKey("DefinitionCount") -and [int]$payload.DefinitionCount -ne $metadataEntries.Count) {
    throw "Canonical DefinitionCount is $($payload.DefinitionCount), but Entries contains $($metadataEntries.Count) metadata entries."
}

# The generated DIC uses Windows line endings because this VSTO add-in is
# built and deployed on Windows. Both outputs are UTF-8 without a BOM.
$dicBytes = $utf8Strict.GetBytes(
    [string]::Join([Environment]::NewLine, $words) + [Environment]::NewLine)
$metadataPayload = [ordered]@{
    Schema = "uzbekorfo-dictionary-metadata-v1"
    DataVersion = if ($payload.ContainsKey("DataVersion")) { $payload.DataVersion } else { $null }
    SourceWordCount = $words.Count
    EntryCount = $metadataEntries.Count
    Entries = $metadataEntries.ToArray()
}
$metadataBytes = $utf8Strict.GetBytes($serializer.Serialize($metadataPayload))

$artifacts = @(
    [pscustomobject]@{ Name = "DIC"; Path = $dicPath; Bytes = $dicBytes },
    [pscustomobject]@{ Name = "metadata index"; Path = $metadataPath; Bytes = $metadataBytes }
)

$staleArtifacts = @($artifacts | Where-Object {
    -not (Test-Path -LiteralPath $_.Path) -or
    -not [System.Linq.Enumerable]::SequenceEqual(
        [byte[]][IO.File]::ReadAllBytes($_.Path), [byte[]]$_.Bytes)
})

if ($Check) {
    if ($staleArtifacts.Count -gt 0) {
        $names = ($staleArtifacts | ForEach-Object { $_.Name }) -join ", "
        throw "Generated dictionary artifact(s) are stale or missing: $names. Run eng\sync-dictionary-data.ps1."
    }

    Write-Host "Generated dictionary artifacts are current: $($words.Count) words, $($metadataEntries.Count) metadata entries."
    return
}

if ($staleArtifacts.Count -eq 0) {
    # The generator itself is an MSBuild input. Touch correct artifacts so a
    # generator-only change does not force every future build.
    foreach ($artifact in $artifacts) {
        [IO.File]::SetLastWriteTimeUtc($artifact.Path, [DateTime]::UtcNow)
    }
    Write-Host "Generated dictionary artifacts are already current: $($words.Count) words, $($metadataEntries.Count) metadata entries."
    return
}

foreach ($artifact in $staleArtifacts) {
    $tempPath = "$($artifact.Path).$([Guid]::NewGuid().ToString('N')).tmp"
    try {
        [IO.File]::WriteAllBytes($tempPath, [byte[]]$artifact.Bytes)
        Move-Item -LiteralPath $tempPath -Destination $artifact.Path -Force
    }
    finally {
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force
        }
    }
}

Write-Host "Generated dictionary artifacts: $($words.Count) words, $($metadataEntries.Count) metadata entries."
