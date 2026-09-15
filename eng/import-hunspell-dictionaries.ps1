<#
.SYNOPSIS
Imports Hunspell root words and retains their affix flags as metadata.

The bundled runtime .dic remains a plain word list. Hunspell's /ABC flags are
not written into that file because the add-in would then treat them as part of
the spelling key. The flags are stored in canonical JSON and its metadata index
for a future flag-aware morphology adapter.
#>
param(
    [string]$LatinPath = (Join-Path $PSScriptRoot "..\uz_UZ.dic"),
    [string]$CyrillicPath = (Join-Path $PSScriptRoot "..\uz_UZ_Cyrl.dic"),
    [string]$DataDirectory = (Join-Path $PSScriptRoot "..\src\UzbekOrfoAddIn\Data"),
    [string]$ReportDirectory = (Join-Path $PSScriptRoot "..\obj\hunspell-import"),
    [switch]$Apply
)

$ErrorActionPreference = "Stop"
[void][Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
$utf8Strict = New-Object Text.UTF8Encoding($false, $true)
$serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$serializer.MaxJsonLength = [int]::MaxValue

function Merge-Flags([string]$left, [string]$right) {
    $characters = New-Object 'System.Collections.Generic.SortedSet[char]'
    foreach ($character in ($left + $right).ToCharArray()) { [void]$characters.Add($character) }
    return -join $characters
}

function Get-LookupKey([string]$value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return "" }
    $result = $value.Trim().Normalize([Text.NormalizationForm]::FormC).ToLowerInvariant()
    $result = $result.Replace([char]0x02BB, "'")
    $result = $result.Replace([char]0x02BC, "'")
    $result = $result.Replace([char]0x2018, "'")
    $result = $result.Replace([char]0x2019, "'")
    return $result
}

function Get-RejectionReason([string]$word) {
    if ([string]::IsNullOrWhiteSpace($word)) { return "blank" }
    if ([regex]::IsMatch($word, "[\p{Cc}\p{Cf}]")) { return "control-or-invisible-character" }
    if ([regex]::IsMatch($word, "\s")) { return "contains-whitespace" }
    if (-not [regex]::IsMatch($word, "^[\p{L}'-]+$")) { return "unsupported-character" }
    if ($word.StartsWith("-") -or $word.EndsWith("-") -or $word.StartsWith("'")) {
        return "invalid-edge-character"
    }
    if ([regex]::IsMatch($word, "[A-Za-z\u00C0-\u024F]") -and
        [regex]::IsMatch($word, "[\u0400-\u052F]")) {
        return "mixed-latin-cyrillic"
    }
    return $null
}

$dictionaryPath = Join-Path $DataDirectory "uzbek_dictionary.json"
foreach ($path in @($LatinPath, $CyrillicPath, $dictionaryPath)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing input file: $path" }
}
$payload = $serializer.DeserializeObject([IO.File]::ReadAllText($dictionaryPath, $utf8Strict))

$sourceWords = @{}
$sourceStats = New-Object 'System.Collections.Generic.List[object]'
$rejections = New-Object 'System.Collections.Generic.List[object]'
foreach ($source in @(
    [pscustomobject]@{ Name = "uz_UZ"; Path = $LatinPath },
    [pscustomobject]@{ Name = "uz_UZ_Cyrl"; Path = $CyrillicPath }
)) {
    $lines = [IO.File]::ReadAllLines((Resolve-Path -LiteralPath $source.Path), $utf8Strict)
    if ($lines.Count -lt 1 -or $lines[0] -notmatch '^\d+$') {
        throw "Hunspell dictionary '$($source.Path)' has no valid entry count."
    }
    $declared = [int]$lines[0]
    $accepted = 0; $duplicates = 0; $rejected = 0; $flagged = 0
    for ($i = 1; $i -lt $lines.Count; $i++) {
        $raw = $lines[$i].Trim()
        if ([string]::IsNullOrWhiteSpace($raw)) { continue }
        $slash = $raw.LastIndexOf('/')
        $wordText = if ($slash -ge 0) { $raw.Substring(0, $slash) } else { $raw }
        $flags = if ($slash -ge 0) { $raw.Substring($slash + 1) } else { "" }
        $word = Get-LookupKey $wordText
        $reason = Get-RejectionReason $word
        # Preserve single-character flags, including special flags. Their
        # meaning belongs to the affix file; metadata alone does not enable it.
        if ($null -eq $reason -and $flags -notmatch '^[\p{L}\p{N}-]*$') {
            $reason = "invalid-affix-flags"
        }
        if ($null -ne $reason) {
            $rejected++
            $rejections.Add([pscustomobject]@{ Source = $source.Name; Line = $i + 1; Reason = $reason; Value = $raw })
            continue
        }
        if ($flags) { $flagged++ }
        $flags = Merge-Flags "" $flags
        if ($sourceWords.ContainsKey($word)) {
            $duplicates++
            $old = $sourceWords[$word]
            if ($flags) {
                $old.Flags = Merge-Flags ([string]$old.Flags) $flags
            }
            continue
        }
        $sourceWords[$word] = [pscustomobject]@{ Word = $word; Flags = $flags; Source = $source.Name }
        $accepted++
    }
    $actualEntries = $lines.Count - 1
    $countMismatch = $declared -ne $actualEntries
    if ($countMismatch) {
        Write-Warning "Hunspell dictionary '$($source.Path)' declares $declared entries but contains $actualEntries."
    }
    $sourceStats.Add([ordered]@{ Source = $source.Name; DeclaredEntries = $declared; Lines = $actualEntries; CountMismatch = $countMismatch; Flagged = $flagged; UniqueAccepted = $accepted; DuplicateEntries = $duplicates; Rejected = $rejected })
}

$canonicalKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
$matched = 0; $updatedFlags = 0; $newEntries = New-Object 'System.Collections.Generic.List[object]'
foreach ($entry in @($payload.Entries)) {
    $key = Get-LookupKey ([string]$entry.Word)
    [void]$canonicalKeys.Add($key)
    if ($sourceWords.ContainsKey($key) -and $sourceWords[$key].Flags) {
        $flags = $sourceWords[$key].Flags
        if (-not $entry.ContainsKey("HunspellFlags") -or $entry.HunspellFlags -cne $flags) {
            $entry["HunspellFlags"] = $flags
            $updatedFlags++
        }
        $matched++
    }
}
foreach ($item in $sourceWords.Values) {
    if ($canonicalKeys.Contains($item.Word)) { continue }
    $newEntry = [ordered]@{ Word = $item.Word; Definition = $null; SpellingRule = $null; GrammarNote = $null; Examples = $null }
    if ($item.Flags) { $newEntry["HunspellFlags"] = $item.Flags }
    $newEntries.Add($newEntry)
}

[IO.Directory]::CreateDirectory($ReportDirectory) | Out-Null
$report = [ordered]@{
    GeneratedAtUtc = [DateTime]::UtcNow.ToString("o")
    Applied = [bool]$Apply
    SourceStats = $sourceStats.ToArray()
    ExistingCanonicalEntries = @($payload.Entries).Count
    MatchedCanonicalWords = $matched
    UpdatedFlagMetadata = $updatedFlags
    MissingCanonicalWords = $newEntries.Count
    Rejected = $rejections.Count
}
$reportPath = Join-Path $ReportDirectory "hunspell-import-summary.json"
[IO.File]::WriteAllText($reportPath, (ConvertTo-Json $report -Depth 6) + [Environment]::NewLine, $utf8Strict)
$missingPath = Join-Path $ReportDirectory "hunspell-missing-canonical-words.tsv"
$missingLines = New-Object 'System.Collections.Generic.List[string]'; $missingLines.Add("Word`tHunspellFlags`tSource")
foreach ($entry in ($newEntries | Sort-Object { $_.Word })) {
    $missingLines.Add("$($entry.Word)`t$($entry.HunspellFlags)`t$($sourceWords[$entry.Word].Source)")
}
[IO.File]::WriteAllLines($missingPath, $missingLines, $utf8Strict)
$rejectionPath = Join-Path $ReportDirectory "hunspell-import-rejections.tsv"
$out = New-Object 'System.Collections.Generic.List[string]'; $out.Add("Source`tLine`tReason`tValue")
foreach ($row in $rejections) { $out.Add("$($row.Source)`t$($row.Line)`t$($row.Reason)`t$($row.Value)") }
[IO.File]::WriteAllLines($rejectionPath, $out, $utf8Strict)

if ($Apply -and ($updatedFlags -gt 0 -or $newEntries.Count -gt 0)) {
    $allEntries = New-Object 'System.Collections.Generic.List[object]'
    foreach ($entry in @($payload.Entries)) { $allEntries.Add($entry) }
    foreach ($entry in ($newEntries | Sort-Object { $_.Word })) { $allEntries.Add($entry) }
    $payload["Entries"] = $allEntries.ToArray()
    $payload["WordCount"] = $allEntries.Count
    $payload["UpdatedAtUtc"] = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
    $date = [DateTime]::UtcNow.ToString("yyyy.MM.dd"); $revision = 1
    if ([string]$payload.DataVersion -match "^$([regex]::Escape($date))\.(\d+)$") { $revision = [int]$Matches[1] + 1 }
    $payload["DataVersion"] = "$date.$revision"
    $temp = "$dictionaryPath.$([Guid]::NewGuid().ToString('N')).tmp"
    try { [IO.File]::WriteAllText($temp, $serializer.Serialize($payload), $utf8Strict); Move-Item -LiteralPath $temp -Destination $dictionaryPath -Force }
    finally { if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force } }
}

Write-Host "Hunspell dictionary review complete."
Write-Host "  Matched canonical words: $matched"
Write-Host "  Updated flag metadata:   $updatedFlags"
Write-Host "  Missing canonical words:  $($newEntries.Count)"
Write-Host "  Rejected:                $($rejections.Count)"
Write-Host "  Report:                  $reportPath"
Write-Host "  Missing words:           $missingPath"
if (-not $Apply) { Write-Host "  Dry run only; use -Apply to update canonical JSON." }
