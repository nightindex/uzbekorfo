<# Generates review material, never linguistic approval. Does not change runtime data. #>
param([string]$OutputDirectory = (Join-Path $PSScriptRoot '..\TestResults\linguist-review'))
$ErrorActionPreference = 'Stop'
$data = Join-Path $PSScriptRoot '..\src\UzbekOrfoAddIn\Data'
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$dictionaryPath = Join-Path $data 'uzbek_dictionary.json'
$rulesPath = Join-Path $data 'uzbek_suffixes.json'
$dictionary = Get-Content -LiteralPath $dictionaryPath -Raw -Encoding UTF8 | ConvertFrom-Json
$rules = Get-Content -LiteralPath $rulesPath -Raw -Encoding UTF8 | ConvertFrom-Json
$issues = [System.Collections.Generic.List[object]]::new()
$seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
$missingLemma = 0; $missingPos = 0; $flagged = 0
foreach ($entry in $dictionary.Entries) {
    $word = [string]$entry.Word
    $reasons = [System.Collections.Generic.List[string]]::new()
    # Screening normalization only; intentionally not claimed to match runtime transliteration.
    $key = $word.Normalize().Trim().ToLowerInvariant() -replace '[\u2018\u2019\u02bb\u02bc\u0060]', "'"
    if (-not $seen.Add($key)) { $reasons.Add('screening-normalized-duplicate') }
    if ($word -match '[a-zA-Z]' -and $word -match '[\u0400-\u04ff]') { $reasons.Add('mixed-latin-cyrillic') }
    if ($word -match '[\p{Cc}\p{Cf}\s]') { $reasons.Add('whitespace-or-invisible-character') }
    if ([string]::IsNullOrWhiteSpace($word)) { $reasons.Add('empty-word') }
    if ([string]::IsNullOrWhiteSpace([string]$entry.Lemma)) { $missingLemma++ }
    if ([string]::IsNullOrWhiteSpace([string]$entry.PartOfSpeech)) { $missingPos++ }
    if ($entry.HunspellFlags) { $flagged++ }
    if ($reasons.Count) {
        $issues.Add([pscustomobject]@{ Word=$word; Reasons=($reasons -join '|'); Lemma=$entry.Lemma;
            PartOfSpeech=$entry.PartOfSpeech; Flags=$entry.HunspellFlags; ReviewerDecision='pending'; Notes='' })
    }
}
$issues | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'dictionary-screening.csv') -NoTypeInformation -Encoding UTF8
$rules.suffixes | Select-Object id,latin,cyrillic,category,family,attachesTo,order,allomorphCondition,
    standaloneOnly,requiredRootFlags,@{n='ReviewerDecision';e={'pending'}},@{n='Evidence';e={''}},@{n='Notes';e={''}} |
    Export-Csv -LiteralPath (Join-Path $OutputDirectory 'suffix-review.csv') -NoTypeInformation -Encoding UTF8
$report = [ordered]@{
    schema='uzbekorfo-linguist-screening-v1'; generatedAtUtc=[DateTime]::UtcNow.ToString('o'); status='engineering-only'
    dictionarySha256=(Get-FileHash -LiteralPath $dictionaryPath -Algorithm SHA256).Hash
    rulesSha256=(Get-FileHash -LiteralPath $rulesPath -Algorithm SHA256).Hash
    entries=@($dictionary.Entries).Count; flaggedEntries=$flagged; missingExplicitLemma=$missingLemma
    missingExplicitPartOfSpeech=$missingPos; screeningIssues=$issues.Count
    suffixes=@($rules.suffixes).Count; importedSuffixes=@($rules.suffixes | Where-Object standaloneOnly).Count
    limitations=@('Screening findings require human adjudication, not automatic deletion.',
        'Missing explicit metadata is not proof that a word is misspelled.',
        'No independent linguist or representative corpus supplied.')
}
$report | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'summary.json') -Encoding UTF8
$report | ConvertTo-Json -Depth 6
