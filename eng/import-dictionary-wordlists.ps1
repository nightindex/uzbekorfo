<#
.SYNOPSIS
Reviews external word lists and imports genuinely missing words.

.DESCRIPTION
Comparison uses the same lookup normalization as TextHelper.NormalizeWord:
trim, lowercase, normalize Uzbek apostrophe variants, and ignore punctuation at
the outside of a token. External records must still be clean single words made
of letters, apostrophes, or internal hyphens. Suspicious records are written to
a TSV report and are never imported automatically.
#>
param(
    [string]$LatinPath = (Join-Path $PSScriptRoot "..\UZ_Latin.txt"),
    [string]$CyrillicPath = (Join-Path $PSScriptRoot "..\UZ_Cyrylic.txt"),
    [string]$DataDirectory = (Join-Path $PSScriptRoot "..\src\UzbekOrfoAddIn\Data"),
    [string]$ReportDirectory = (Join-Path $PSScriptRoot "..\obj\dictionary-import"),
    [switch]$Apply
)

$ErrorActionPreference = "Stop"

[void][Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
$utf8Strict = New-Object Text.UTF8Encoding($false, $true)
$serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$serializer.MaxJsonLength = [int]::MaxValue

function Get-LookupKey([string]$value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return "" }

    $result = $value.Trim().Normalize([Text.NormalizationForm]::FormC).ToLowerInvariant()
    $result = $result.Replace([char]0x02BB, "'")
    $result = $result.Replace([char]0x02BC, "'")
    $result = $result.Replace([char]0x2018, "'")
    $result = $result.Replace([char]0x2019, "'")
    $result = [regex]::Replace($result, "^[^\p{L}']+", "")
    $result = [regex]::Replace($result, "[^\p{L}']+$", "")
    return $result
}

function Get-RejectionReason([string]$raw, [string]$word) {
    if ([string]::IsNullOrWhiteSpace($raw)) { return "blank" }
    if ($raw.Trim().StartsWith("#")) { return "comment" }
    if ([regex]::IsMatch($raw, "[\p{Cc}\p{Cf}]")) { return "control-or-invisible-character" }
    if ([string]::IsNullOrWhiteSpace($word) -or -not [regex]::IsMatch($word, "\p{L}")) {
        return "no-letter"
    }
    if ([regex]::IsMatch($word, "\s")) { return "contains-whitespace" }
    if (-not [regex]::IsMatch($word, "^[\p{L}'-]+$")) { return "unsupported-character" }
    # An apostrophe may be the last code point of an Uzbek letter such as o' or
    # g' (for example, obro' and tog'). It is only invalid at the beginning.
    if ($word.StartsWith("-") -or $word.EndsWith("-") -or $word.StartsWith("'")) {
        return "invalid-edge-character"
    }

    $hasLatin = [regex]::IsMatch($word, "[A-Za-z\u00C0-\u024F]")
    $hasCyrillic = [regex]::IsMatch($word, "[\u0400-\u052F]")
    if ($hasLatin -and $hasCyrillic) { return "mixed-latin-cyrillic" }
    return $null
}

$dictionaryPath = Join-Path $DataDirectory "uzbek_dictionary.json"
foreach ($path in @($LatinPath, $CyrillicPath, $dictionaryPath)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Missing input file: $path"
    }
}

try {
    $payload = $serializer.DeserializeObject([IO.File]::ReadAllText($dictionaryPath, $utf8Strict))
}
catch {
    throw "Invalid UTF-8 or JSON in '$dictionaryPath': $($_.Exception.Message)"
}

$existingKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
$existingFirstWords = @{}
$existingDuplicates = New-Object 'System.Collections.Generic.List[object]'
foreach ($entry in @($payload.Entries)) {
    $existingWord = [string]$entry.Word
    $existingKey = Get-LookupKey $existingWord
    if (-not $existingKeys.Add($existingKey)) {
        $existingDuplicates.Add([pscustomobject]@{
            LookupKey = $existingKey
            FirstWord = [string]$existingFirstWords[$existingKey]
            DuplicateWord = $existingWord
        })
    }
    else {
        $existingFirstWords[$existingKey] = $existingWord
    }
}

$candidateKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
$missingEntries = New-Object 'System.Collections.Generic.List[object]'
$rejections = New-Object 'System.Collections.Generic.List[object]'
$sourceSummaries = New-Object 'System.Collections.Generic.List[object]'

$sources = @(
    [pscustomobject]@{ Name = "UZ_Latin"; Path = $LatinPath },
    [pscustomobject]@{ Name = "UZ_Cyrylic"; Path = $CyrillicPath }
)

foreach ($source in $sources) {
    $lineNumber = 0
    $accepted = 0
    $alreadyPresent = 0
    $duplicateCandidate = 0
    $rejected = 0

    try {
        $lines = [IO.File]::ReadLines((Resolve-Path -LiteralPath $source.Path), $utf8Strict)
        foreach ($rawLine in $lines) {
            $lineNumber++
            $word = Get-LookupKey $rawLine
            $reason = Get-RejectionReason $rawLine $word
            if ($null -ne $reason) {
                $rejected++
                $rejections.Add([pscustomobject]@{
                    Source = $source.Name
                    Line = $lineNumber
                    Reason = $reason
                    Value = $rawLine
                })
                continue
            }

            if ($existingKeys.Contains($word)) {
                $alreadyPresent++
                continue
            }
            if (-not $candidateKeys.Add($word)) {
                $duplicateCandidate++
                continue
            }

            $accepted++
            $missingEntries.Add([ordered]@{
                Word = $word
                Definition = $null
                SpellingRule = $null
                GrammarNote = $null
                Examples = $null
            })
        }
    }
    catch {
        throw "Invalid UTF-8 or unreadable source '$($source.Path)': $($_.Exception.Message)"
    }

    $sourceSummaries.Add([ordered]@{
        Source = $source.Name
        Path = [IO.Path]::GetFullPath($source.Path)
        Lines = $lineNumber
        MissingAccepted = $accepted
        AlreadyPresent = $alreadyPresent
        DuplicateCandidate = $duplicateCandidate
        Rejected = $rejected
    })
}

[IO.Directory]::CreateDirectory($ReportDirectory) | Out-Null
$summaryPath = Join-Path $ReportDirectory "dictionary-import-summary.json"
$rejectionsPath = Join-Path $ReportDirectory "dictionary-import-rejections.tsv"
$duplicatesPath = Join-Path $ReportDirectory "dictionary-existing-normalized-duplicates.tsv"
$summary = [ordered]@{
    GeneratedAtUtc = [DateTime]::UtcNow.ToString("o")
    Applied = [bool]$Apply
    ExistingDictionaryEntries = @($payload.Entries).Count
    ExistingLookupKeys = $existingKeys.Count
    ExistingNormalizedDuplicates = $existingDuplicates.Count
    MissingAccepted = $missingEntries.Count
    Rejected = $rejections.Count
    Sources = $sourceSummaries.ToArray()
}
[IO.File]::WriteAllText(
    $summaryPath,
    (ConvertTo-Json -InputObject $summary -Depth 5) + [Environment]::NewLine,
    $utf8Strict)

$tsvLines = New-Object 'System.Collections.Generic.List[string]'
$tsvLines.Add("Source`tLine`tReason`tValue")
foreach ($row in $rejections) {
    $safeValue = ([string]$row.Value).Replace("`t", " ").Replace("`r", " ").Replace("`n", " ")
    $tsvLines.Add("$($row.Source)`t$($row.Line)`t$($row.Reason)`t$safeValue")
}
[IO.File]::WriteAllLines($rejectionsPath, $tsvLines, $utf8Strict)

$duplicateLines = New-Object 'System.Collections.Generic.List[string]'
$duplicateLines.Add("LookupKey`tFirstWord`tDuplicateWord")
foreach ($row in $existingDuplicates) {
    $duplicateLines.Add("$($row.LookupKey)`t$($row.FirstWord)`t$($row.DuplicateWord)")
}
[IO.File]::WriteAllLines($duplicatesPath, $duplicateLines, $utf8Strict)

if ($Apply -and $missingEntries.Count -gt 0) {
    $allEntries = New-Object 'System.Collections.Generic.List[object]'
    foreach ($entry in @($payload.Entries)) { $allEntries.Add($entry) }
    foreach ($entry in ($missingEntries | Sort-Object { $_.Word })) { $allEntries.Add($entry) }

    $payload["Entries"] = $allEntries.ToArray()
    $payload["WordCount"] = $allEntries.Count
    $payload["UpdatedAtUtc"] = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")

    $versionDate = [DateTime]::UtcNow.ToString("yyyy.MM.dd")
    $nextRevision = 1
    if ([string]$payload.DataVersion -match "^$([regex]::Escape($versionDate))\.(\d+)$") {
        $nextRevision = [int]$Matches[1] + 1
    }
    $payload["DataVersion"] = "$versionDate.$nextRevision"

    $tempPath = "$dictionaryPath.$([Guid]::NewGuid().ToString('N')).tmp"
    try {
        [IO.File]::WriteAllText($tempPath, $serializer.Serialize($payload), $utf8Strict)
        [void]$serializer.DeserializeObject([IO.File]::ReadAllText($tempPath, $utf8Strict))
        Move-Item -LiteralPath $tempPath -Destination $dictionaryPath -Force
    }
    finally {
        if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force }
    }
}

Write-Host "Dictionary review complete."
Write-Host "  Existing lookup keys: $($existingKeys.Count)"
Write-Host "  Missing accepted:     $($missingEntries.Count)"
Write-Host "  Rejected for review:  $($rejections.Count)"
Write-Host "  Summary:              $summaryPath"
Write-Host "  Rejections:           $rejectionsPath"
Write-Host "  Existing duplicates:  $duplicatesPath"
if (-not $Apply) { Write-Host "  Dry run only; use -Apply to update the canonical dictionary." }
