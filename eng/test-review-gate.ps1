<# Isolated synthetic tests of approval validation, not actual reviewer sign-off. #>
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$fixture = Join-Path $root ('TestResults\review-gate-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fixture -Force | Out-Null
$corpusPath = Join-Path $fixture 'corpus.json'
$approvalPath = Join-Path $fixture 'approval.json'
$corpus = Get-Content -Raw -Encoding UTF8 (Join-Path $root 'tests\Corpora\uzbek_morphology_corpus.json') | ConvertFrom-Json
$corpus.review.linguistStatus = 'approved'
$corpus.review.linguistName = 'SYNTHETIC TEST - NOT A REVIEWER'
$corpus.review.reviewedAtUtc = '2026-01-01T00:00:00Z'
$corpus | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $corpusPath -Encoding UTF8
$approval = [ordered]@{
    schema='uzbekorfo-linguist-approval-v1'; status='approved'; independentReviewer=$true
    reviewer='SYNTHETIC TEST - NOT A REVIEWER'; evidence='synthetic fixture only'; reviewedAtUtc='2026-01-01T00:00:00Z'
    parserCorpusSha256=(Get-FileHash -LiteralPath $corpusPath).Hash
    rulesSha256=(Get-FileHash -LiteralPath (Join-Path $root 'src\UzbekOrfoAddIn\Data\uzbek_suffixes.json')).Hash
    dictionarySha256=(Get-FileHash -LiteralPath (Join-Path $root 'src\UzbekOrfoAddIn\Data\uzbek_dictionary.json')).Hash
    evaluationCorpusSha256=(Get-FileHash -LiteralPath (Join-Path $root 'tests\Corpora\spelling_engineering_seed.json')).Hash
}
function Test-Gate([string]$Name, [bool]$ShouldPass) {
    $approval | ConvertTo-Json | Set-Content -LiteralPath $approvalPath -Encoding UTF8
    $passed = $true
    try { & (Join-Path $PSScriptRoot 'validate-linguist-review.ps1') -CorpusPath $corpusPath -ApprovalPath $approvalPath -RequireApproved 6>$null }
    catch { $passed = $false }
    if ($passed -ne $ShouldPass) { throw "Unexpected approval validation result: $Name" }
    Write-Host "PASS: $Name"
}
Test-Gate 'matching synthetic snapshot accepted' $true
foreach ($key in @('rulesSha256','dictionarySha256','parserCorpusSha256','evaluationCorpusSha256')) {
    $original = $approval[$key]; $approval[$key] = ('0' * 64)
    Test-Gate "stale $key rejected" $false
    $approval[$key] = $original
}
$approval.independentReviewer = $false
Test-Gate 'non-independent approval rejected' $false
$approval.independentReviewer = $true
$approval.reviewedAtUtc = 'invalid-date'
Test-Gate 'invalid review date rejected' $false
$approval.reviewedAtUtc = '2026-01-01T00:00:00Z'
$approval.status = 'pending'
Test-Gate 'pending approval rejected' $false
