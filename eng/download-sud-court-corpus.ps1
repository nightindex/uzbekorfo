[CmdletBinding()]
param(
    [string]$RepositoryRoot = (Join-Path $PSScriptRoot '..'),
    [string]$CyrillicRoot,
    [string]$LatinRoot,
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$RepositoryRoot = [IO.Path]::GetFullPath($RepositoryRoot)
if ([string]::IsNullOrWhiteSpace($CyrillicRoot)) { $CyrillicRoot = Join-Path $RepositoryRoot 'corpus/raw_cyrillic/sud-court' }
if ([string]::IsNullOrWhiteSpace($LatinRoot)) { $LatinRoot = Join-Path $RepositoryRoot 'corpus/raw-latin/sud-court' }
$CyrillicRoot = [IO.Path]::GetFullPath($CyrillicRoot)
$LatinRoot = [IO.Path]::GetFullPath($LatinRoot)

$pages = @(
    'https://old.sud.uz/%D1%81%D1%83%D0%B4-%D1%84%D0%B0%D0%BE%D0%BB%D0%B8%D1%8F%D1%82%D0%B8%D0%B3%D0%B0-%D1%82%D0%B5%D0%B3%D0%B8%D1%88%D0%BB%D0%B8/',
    'https://old.sud.uz/%D1%84%D1%83%D2%9B%D0%B0%D1%80%D0%BE%D0%BB%D0%B8%D0%BA-%D0%B8%D1%88%D0%BB%D0%B0%D1%80%D0%B8-%D0%B1%D1%9E%D0%B9%D0%B8%D1%87%D0%B0-%D1%82%D1%83%D1%88%D1%83%D0%BD%D1%82%D0%B8%D1%80%D0%B8%D1%88%D0%BB/',
    'https://old.sud.uz/%D0%BC%D0%B0%D1%8A%D0%BC%D1%83%D1%80%D0%B8%D0%B9-%D0%B8%D1%88%D0%BB%D0%B0%D1%80-%D0%B1%D1%1E%D0%B9%D0%B8%D1%87%D0%B0-%D1%82%D1%83%D1%88%D1%83%D0%BD%D1%82%D0%B8%D1%80%D0%B8%D1%88%D0%BB%D0%B0%D1%80/',
    'https://old.sud.uz/%D0%B8%D2%9B%D1%82%D0%B8%D1%81%D0%BE%D0%B4%D0%B8%D0%B9-%D0%B8%D1%88%D0%BB%D0%B0%D1%80%D0%B8-%D0%B1%D1%9E%D0%B9%D0%B8%D1%87%D0%B0-%D1%82%D1%83%D1%88%D1%83%D0%BD%D1%82%D0%B8%D1%80%D0%B8%D1%88%D0%BB%D0%B0/',
    'https://old.sud.uz/%D1%81%D1%83%D0%B4-%D1%84%D0%B0%D0%BE%D0%BB%D0%B8%D1%8F%D1%82%D0%B8%D0%BD%D0%B8%D0%BD%D0%B3-%D1%83%D0%BC%D1%83%D0%BC%D0%B8%D0%B9-%D0%BC%D0%B0%D1%81%D0%B0%D0%BB%D0%B0%D0%BB%D0%B0%D1%80%D0%B8/'
)
# Normalize the URL-encoded Cyrillic "ў" spelling in case an editor rewrites it.
$pages = $pages | ForEach-Object { $_ -replace '%D1%1E', '%D1%9E' }
$pageFolders = @(
    'sud-faoliyatiga-tegishli',
    'fuqarolik-ishlari-tushuntirishlari',
    'mamuriy-ishlar-tushuntirishlari',
    'iqtisodiy-ishlar-tushuntirishlari',
    'sud-faoliyatining-umumiy-masalalari'
)

function Get-Ids([string]$html) {
    $ids = [Collections.Generic.HashSet[string]]::new()
    foreach ($match in [regex]::Matches($html, '(?i)(?:https?://(?:www\.)?lex\.uz)?/docs/(-?\d+)')) {
        $id = $match.Groups[1].Value
        if ($id.StartsWith('-')) { $id = $id.Substring(1) }
        if ($id -match '^\d+$') { [void]$ids.Add($id) }
    }
    return @($ids | Sort-Object {[int64]$_})
}

function Test-WordExport([string]$path, [bool]$latin) {
    if (-not (Test-Path -LiteralPath $path)) { return $false }
    $bytes = [IO.File]::ReadAllBytes($path)
    if ($bytes.Length -lt 1000) { return $false }
    $html = [Text.Encoding]::UTF8.GetString($bytes)
    if ($html -notmatch '(?i)<html') { return $false }
    # Court pages can include Uzbek and Russian documents. Preserve every valid
    # official Word export; script filtering belongs to corpus build/evaluation,
    # not to acquisition.
    if ($html -match '(?i)Page not found') { return $false }
    return $true
}

$all = [Collections.Generic.HashSet[string]]::new()
$idFolders = @{}
$pageResults = [Collections.Generic.List[object]]::new()
for ($pageIndex = 0; $pageIndex -lt $pages.Count; $pageIndex++) {
    $page = $pages[$pageIndex]
    Write-Host "Reading $page"
    try {
        $response = Invoke-WebRequest -Uri $page -MaximumRedirection 5 -TimeoutSec 90
        $found = @(Get-Ids $response.Content)
        foreach ($id in $found) {
            [void]$all.Add($id)
            if (-not $idFolders.ContainsKey($id)) { $idFolders[$id] = $pageFolders[$pageIndex] }
        }
        $pageResults.Add([ordered]@{ url = $page; status = 'ok'; documentsDiscovered = $found.Count })
    }
    catch {
        Write-Warning "Page unavailable; continuing without it: $page ($($_.Exception.Message))"
        $pageResults.Add([ordered]@{ url = $page; status = 'unavailable'; error = $_.Exception.Message })
    }
}
$ids = @($all | Sort-Object {[int64]$_})
if ($ids.Count -eq 0) { throw 'No LexUZ document links were discovered.' }

function Download-Variant([string]$root, [bool]$latin) {
    New-Item -ItemType Directory -Force -Path $root | Out-Null
    $label = if ($latin) { 'Latin' } else { 'Cyrillic' }
    $rows = [Collections.Generic.List[object]]::new()
    $downloaded = 0; $skipped = 0
    for ($i = 0; $i -lt $ids.Count; $i++) {
        $id = $ids[$i]
        $relative = Join-Path $idFolders[$id] "lexuz-$id.doc"
        $destination = Join-Path $root $relative
        $legacyDestination = Join-Path $root "lexuz-$id.doc"
        $temporary = "$destination.download"
        $exportId = if ($latin) { "-$id" } else { $id }
        $url = "https://lex.uz/docs/$exportId`?type=doc"
        if (-not $Force -and -not (Test-Path -LiteralPath $destination) -and
            (Test-WordExport $legacyDestination $latin)) {
            New-Item -ItemType Directory -Force -Path (Split-Path -Parent $destination) | Out-Null
            Move-Item -LiteralPath $legacyDestination -Destination $destination
            $skipped++; Write-Host "[$($i + 1)/$($ids.Count)] Organized $relative"
            $rows.Add([ordered]@{ id = $id; path = $relative.Replace('\\', '/'); url = $url; status = 'organized' })
        } elseif ((Test-WordExport $destination $latin) -and -not $Force) {
            $skipped++; Write-Host "[$($i + 1)/$($ids.Count)] Kept $relative"
            $rows.Add([ordered]@{ id = $id; path = $relative.Replace('\\', '/'); url = $url; status = 'kept' })
        } else {
            if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force }
            Write-Host "[$($i + 1)/$($ids.Count)] Downloading $relative ($label)"
            try {
                Invoke-WebRequest -Uri $url -OutFile $temporary -MaximumRedirection 5 -TimeoutSec 90
                if (-not (Test-WordExport $temporary $latin)) { throw 'Response failed script/Word-export validation.' }
                Move-Item -LiteralPath $temporary -Destination $destination -Force
                $downloaded++
            } catch {
                if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force }
                Write-Warning "Document unavailable; continuing: $url ($($_.Exception.Message))"
                $rows.Add([ordered]@{ id = $id; path = $relative; url = $url; status = 'unavailable'; error = $_.Exception.Message })
                continue
            }
            $rows.Add([ordered]@{ id = $id; path = $relative.Replace('\\', '/'); url = $url; status = 'downloaded' })
        }
    }
    $manifest = [ordered]@{
        schema = 'matnai-sud-court-corpus-v1'; script = if ($latin) { 'Latin' } else { 'Cyrillic' }
        downloadedUtc = [DateTime]::UtcNow.ToString('o'); sourcePages = $pages; pageResults = @($pageResults)
        documents = @($rows)
    }
    [IO.File]::WriteAllText((Join-Path $root 'manifest.json'), ($manifest | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding($false)))
    Write-Host "Completed $label`: $downloaded downloaded, $skipped kept, $($ids.Count) total."
}

Download-Variant $CyrillicRoot $false
Download-Variant $LatinRoot $true
