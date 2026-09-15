param(
    [string]$Destination = (Join-Path $PSScriptRoot '..\obj\restored-hunspell'),
    [string]$BundlePath = (Join-Path $PSScriptRoot '..\docs\reference\hunspell_sources.json')
)
$ErrorActionPreference = 'Stop'
[void][Reflection.Assembly]::LoadWithPartialName('System.Web.Extensions')
$serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$serializer.MaxJsonLength = [int]::MaxValue
$bundle = $serializer.DeserializeObject([IO.File]::ReadAllText($BundlePath))
if ($bundle.Schema -ne 'uzbekorfo-hunspell-v1') { throw 'Unsupported bundle schema.' }
$destinationPath = [IO.Path]::GetFullPath($Destination)
[void][IO.Directory]::CreateDirectory($destinationPath)
foreach ($file in $bundle.Files) {
    if ([IO.Path]::GetFileName($file.Name) -cne $file.Name) { throw 'Invalid source filename.' }
    $path = Join-Path $destinationPath $file.Name
    if (Test-Path -LiteralPath $path) { throw "Refusing to overwrite $path" }
    $bytes = [Convert]::FromBase64String($file.Base64)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $hash = [BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '') }
    finally { $sha.Dispose() }
    if ($hash -cne $file.Sha256) { throw "Source checksum mismatch: $($file.Name)" }
    [IO.File]::WriteAllBytes($path, $bytes)
}
Write-Host "Restored original sources to $destinationPath"
