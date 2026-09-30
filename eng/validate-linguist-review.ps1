<#
.SYNOPSIS
Validates morphology corpus review metadata and optionally enforces linguist approval.
#>
param(
    [string]$CorpusPath = (Join-Path $PSScriptRoot '..\tests\Corpora\uzbek_morphology_corpus.json'),
    [string]$RulesPath = (Join-Path $PSScriptRoot '..\src\UzbekOrfoAddIn\Data\uzbek_suffixes.json'),
    [string]$DictionaryPath = (Join-Path $PSScriptRoot '..\src\UzbekOrfoAddIn\Data\uzbek_dictionary.json'),
    [string]$EvaluationCorpusPath = (Join-Path $PSScriptRoot '..\tests\Corpora\spelling_engineering_seed.json'),
    [string]$ApprovalPath = (Join-Path $PSScriptRoot '..\docs\reviews\linguist-approval.json'),
    [switch]$RequireApproved
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath $CorpusPath)) { throw "Morphology corpus is missing: $CorpusPath" }
if (-not (Test-Path -LiteralPath $RulesPath)) { throw "Morphology rules are missing: $RulesPath" }

$corpus = Get-Content -LiteralPath $CorpusPath -Raw -Encoding UTF8 | ConvertFrom-Json
$rules = Get-Content -LiteralPath $RulesPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ($corpus.schema -ne 'uzbekorfo-morphology-corpus-v1') { throw 'Unsupported morphology corpus schema.' }
if ($corpus.ruleVersion -ne $rules.version) { throw "Corpus ruleVersion '$($corpus.ruleVersion)' does not match '$($rules.version)'." }
if ($corpus.review.engineeringStatus -ne 'reviewed') { throw 'The corpus has not passed engineering review.' }
if (@($corpus.cases).Count -lt 15) { throw 'The morphology corpus must contain at least 15 cases.' }
if (@($corpus.cases | Where-Object expectedCorrect).Count -eq 0 -or
    @($corpus.cases | Where-Object { -not $_.expectedCorrect }).Count -eq 0) {
    throw 'The morphology corpus must contain both correct and incorrect forms.'
}

$approved = $corpus.review.linguistStatus -eq 'approved' -and
    -not [string]::IsNullOrWhiteSpace([string]$corpus.review.linguistName) -and
    -not [string]::IsNullOrWhiteSpace([string]$corpus.review.reviewedAtUtc)

# A schema version is not a content version. Bind independent sign-off to the exact files.
$snapshotApproved = $false
if (Test-Path -LiteralPath $ApprovalPath) {
    $approval = Get-Content -LiteralPath $ApprovalPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($approval.status -eq 'approved') {
        if ($approval.schema -ne 'uzbekorfo-linguist-approval-v1' -or $approval.independentReviewer -ne $true -or
            [string]::IsNullOrWhiteSpace([string]$approval.reviewer) -or
            [string]::IsNullOrWhiteSpace([string]$approval.evidence)) {
            throw 'Independent approval requires schema, reviewer identity, independence declaration and evidence reference.'
        }
        $reviewedAt = [DateTimeOffset]::MinValue
        if (-not [DateTimeOffset]::TryParse([string]$approval.reviewedAtUtc, [ref]$reviewedAt) -or
            $reviewedAt -gt [DateTimeOffset]::UtcNow) { throw 'Invalid approval date.' }
        $inputs = @{
            rulesSha256=$RulesPath; dictionarySha256=$DictionaryPath
            parserCorpusSha256=$CorpusPath; evaluationCorpusSha256=$EvaluationCorpusPath
        }
        foreach ($key in $inputs.Keys) {
            $expected = [string]$approval.$key
            if ($expected -notmatch '^[0-9a-fA-F]{64}$' -or
                $expected -ne (Get-FileHash -LiteralPath $inputs[$key] -Algorithm SHA256).Hash) {
                throw "Linguistic approval is stale or incomplete: $key. Re-review the changed data."
            }
        }
        $snapshotApproved = $true
    }
}
$approved = $approved -and $snapshotApproved

if ($RequireApproved -and -not $approved) {
    throw 'Production release is blocked: an Uzbek linguist must approve the suffix rules and corpus review metadata.'
}
if ($approved) {
    Write-Host "[PASS] Linguist review approved by $($corpus.review.linguistName) at $($corpus.review.reviewedAtUtc)."
} else {
    Write-Warning 'Uzbek linguist review is pending. Engineering tests may run, but production release is blocked.'
}
