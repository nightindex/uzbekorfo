param([string]$Root = (Split-Path $PSScriptRoot -Parent), [switch]$RemoveSources)
$ErrorActionPreference = 'Stop'
[void][Reflection.Assembly]::LoadWithPartialName('System.Web.Extensions')
$serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$serializer.MaxJsonLength = [int]::MaxValue
$files = foreach ($name in @('uz_UZ.dic','uz_UZ_Cyrl.dic','suf_uz_UZ.txt','suf_uz_UZ_Cyrl.txt','UZ_Latin.txt','UZ_Cyrylic.txt')) {
    $path = Join-Path $Root $name
    $bytes = [IO.File]::ReadAllBytes($path)
    [ordered]@{ Name=$name; Sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash; Base64=[Convert]::ToBase64String($bytes) }
}
$bundle = [ordered]@{ Schema='uzbekorfo-hunspell-v1'; Files=@($files) }
$target = Join-Path $Root 'docs\reference\hunspell_sources.json'
[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($target))
[IO.File]::WriteAllText($target, $serializer.Serialize($bundle), (New-Object Text.UTF8Encoding($false)))
$roundTrip = $serializer.DeserializeObject([IO.File]::ReadAllText($target))
foreach ($file in $roundTrip.Files) {
    $original = [IO.File]::ReadAllBytes((Join-Path $Root $file.Name))
    $restored = [Convert]::FromBase64String($file.Base64)
    if (-not [System.Linq.Enumerable]::SequenceEqual([byte[]]$original, [byte[]]$restored)) { throw "Round-trip failed: $($file.Name)" }
}
Write-Host 'All six sources migrated and verified byte for byte.'
if ($RemoveSources) {
    foreach ($file in $roundTrip.Files) {
        $sourcePath = [IO.Path]::GetFullPath((Join-Path $Root $file.Name))
        if ([IO.Path]::GetDirectoryName($sourcePath) -ne [IO.Path]::GetFullPath($Root).TrimEnd('\')) {
            throw "Source outside project root: $sourcePath"
        }
        Remove-Item -LiteralPath $sourcePath
    }
    Write-Host 'Removed six redundant source files. All are recoverable from docs/reference/hunspell_sources.json.'
}
