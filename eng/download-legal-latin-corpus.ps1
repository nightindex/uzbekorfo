[CmdletBinding()]
param(
    [string]$RepositoryRoot = (Join-Path $PSScriptRoot '..'),
    [string]$DownloadRoot,
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$RepositoryRoot = [IO.Path]::GetFullPath($RepositoryRoot)
if ([string]::IsNullOrWhiteSpace($DownloadRoot)) {
    $DownloadRoot = Join-Path $RepositoryRoot 'corpus/raw-latin'
}
$DownloadRoot = [IO.Path]::GetFullPath($DownloadRoot)
$reportPath = Join-Path $RepositoryRoot 'docs/prediction/corpus-report.json'
if (-not (Test-Path -LiteralPath $reportPath)) {
    throw "Corpus report was not found: $reportPath"
}

function Test-LatinWordExport([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    $bytes = [IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -lt 1000) { return $false }
    $text = [Text.Encoding]::UTF8.GetString($bytes)
    if ($text -notmatch '<html|<HTML') { return $false }
    $plain = [System.Net.WebUtility]::HtmlDecode([regex]::Replace($text, '<[^>]+>', ' '))
    $latin = [regex]::Matches($plain, '[A-Za-z]').Count
    $cyrillic = [regex]::Matches($plain, '\p{IsCyrillic}').Count
    return $latin -gt 100 -and $latin -gt ($cyrillic * 2)
}

$report = Get-Content -LiteralPath $reportPath -Raw | ConvertFrom-Json
$sources = @(
    $report.files |
        ForEach-Object {
            $match = [regex]::Match([string]$_.url, '/docs/(\d+)\?type=doc$')
            if (-not $match.Success) {
                throw "Cannot determine a Lex.uz document ID from '$($_.url)'."
            }
            $source = ([string]$_.source).Replace('/', '\\')
            $prefix = "corpus\\raw\\"
            if (-not $source.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
                throw "Unexpected corpus source path '$source'."
            }
            [pscustomobject]@{
                Id = $match.Groups[1].Value
                RelativePath = $source.Substring($prefix.Length)
                Url = "https://lex.uz/docs/-$($match.Groups[1].Value)?type=doc"
            }
        } |
        Sort-Object RelativePath
)
if ($sources.Count -ne 57) {
    throw "Expected 57 public legal sources from the checked-in corpus report; found $($sources.Count)."
}

New-Item -ItemType Directory -Force -Path $DownloadRoot | Out-Null
$downloaded = 0
$skipped = 0
for ($index = 0; $index -lt $sources.Count; $index++) {
    $source = $sources[$index]
    $destination = Join-Path $DownloadRoot $source.RelativePath
    $temporary = "$destination.download"
    if ((Test-LatinWordExport $destination) -and -not $Force) {
        $skipped++
        Write-Host "[$($index + 1)/$($sources.Count)] Kept $($source.RelativePath)"
        continue
    }
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force }
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $destination) | Out-Null
    Write-Host "[$($index + 1)/$($sources.Count)] Downloading $($source.RelativePath)"
    try {
        Invoke-WebRequest -Uri $source.Url -OutFile $temporary -MaximumRedirection 5 -TimeoutSec 90
        if (-not (Test-LatinWordExport $temporary)) {
            throw 'The response was not a valid Uzbek Latin Word-compatible HTML export.'
        }
        Move-Item -LiteralPath $temporary -Destination $destination -Force
        $downloaded++
    }
    catch {
        if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force }
        throw "Failed to download $($source.Url): $($_.Exception.Message)"
    }
}

$manifest = [ordered]@{
    schema = 'matnai-legal-latin-corpus-v1'
    downloadedUtc = [DateTime]::UtcNow.ToString('o')
    sourcePages = @(
        'https://old.sud.uz/%D0%BA%D0%BE%D0%B4%D0%B5%D0%BA%D1%81%D0%BB%D0%B0%D1%80/',
        'https://old.sud.uz/%D2%9B%D0%BE%D0%BD%D1%83%D0%BD%D1%87%D0%B8%D0%BB%D0%B8%D0%BA/',
        'https://old.sud.uz/%D2%9B%D0%BE%D0%BD%D1%83%D0%BD%D0%BB%D0%B0%D1%80/',
        'https://old.sud.uz/%D0%BF%D1%80%D0%B5%D0%B7%D0%B8%D0%B4%D0%B5%D0%BD%D1%82-%D1%84%D0%B0%D1%80%D0%BC%D0%BE%D0%BD%D0%BB%D0%B0%D1%80%D0%B8-%D0%B2%D0%B0-%D2%9B%D0%B0%D1%80%D0%BE%D1%80%D0%BB%D0%B0%D1%80%D0%B8/',
        'https://old.sud.uz/%D0%B2%D0%B0%D0%B7%D0%B8%D1%80%D0%BB%D0%B0%D1%80-%D0%BC%D0%B0%D2%B3%D0%BA%D0%B0%D0%BC%D0%B0%D1%81%D0%B8-%D2%9B%D0%B0%D1%80%D0%BE%D1%80%D0%BB%D0%B0%D1%80%D0%B8/'
    )
    documents = @($sources | ForEach-Object { [ordered]@{ id = $_.Id; path = $_.RelativePath.Replace('\\', '/'); url = $_.Url } })
}
[IO.File]::WriteAllText((Join-Path $DownloadRoot 'manifest.json'), ($manifest | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding($false)))
Write-Host "Completed: $downloaded downloaded, $skipped already valid, $($sources.Count) total."
