param(
    [ValidateSet('Offline', 'Web')]
    [string]$Channel = 'Offline',
    [string]$UpdateUrl,
    [string]$Configuration = 'Release',
    [string]$PublishDirectory,
    [string]$ManifestKeyFile = $env:UZBEKORFO_MANIFEST_KEY_FILE,
    [string]$ManifestCertificateThumbprint = $env:UZBEKORFO_MANIFEST_CERTIFICATE_THUMBPRINT,
    [switch]$SkipUnitTests,
    [switch]$SkipUiTests,
    [switch]$SkipWordTests,
    [switch]$SkipPublish,
    [switch]$AllowNonReleaseBranch,
    [switch]$AllowDirtyWorktree,
    [switch]$AllowUntrustedCertificate
)

$ErrorActionPreference = 'Stop'

function Assert-LastExitCode([string]$Description) {
    if ($LASTEXITCODE -ne 0) {
        throw "$Description failed with exit code $LASTEXITCODE."
    }
}

function Get-SigningConfiguration(
    [string]$RepositoryRoot,
    [string]$KeyFile,
    [string]$Thumbprint,
    [bool]$AllowUntrustedCertificate) {
    if ([string]::IsNullOrWhiteSpace($KeyFile) -and [string]::IsNullOrWhiteSpace($Thumbprint)) {
        $propsPath = Join-Path $RepositoryRoot 'Directory.Build.props'
        if (-not (Test-Path -LiteralPath $propsPath)) {
            throw 'No signing configuration found. Configure the ignored Directory.Build.props file or secure CI environment variables.'
        }

        [xml]$props = Get-Content -Raw -LiteralPath $propsPath
        $keyNode = $props.SelectSingleNode("//*[local-name()='ManifestKeyFile']")
        $thumbprintNode = $props.SelectSingleNode("//*[local-name()='ManifestCertificateThumbprint']")
        $KeyFile = if ($keyNode) { $keyNode.InnerText.Trim() } else { '' }
        $Thumbprint = if ($thumbprintNode) { $thumbprintNode.InnerText.Trim() } else { '' }
    }

    if ([string]::IsNullOrWhiteSpace($KeyFile) -or [string]::IsNullOrWhiteSpace($Thumbprint)) {
        throw 'Set both ManifestKeyFile and ManifestCertificateThumbprint through local props or secure CI environment variables.'
    }
    $usesUntrustedCertificate =
        $Thumbprint -match 'REPLACE|TEMPORARY|EXAMPLE' -or
        $KeyFile -match '(?i)temporary|example|test'
    if ($usesUntrustedCertificate -and -not $AllowUntrustedCertificate) {
        throw 'Release signing configuration still references a placeholder or temporary certificate. Use -AllowUntrustedCertificate only for an explicitly accepted local-certificate release.'
    }

    $KeyFile = $KeyFile.Replace('$(MSBuildThisFileDirectory)', "$RepositoryRoot\")
    if (-not [System.IO.Path]::IsPathRooted($KeyFile)) {
        $KeyFile = Join-Path $RepositoryRoot $KeyFile
    }
    if (-not (Test-Path -LiteralPath $KeyFile)) {
        throw "ManifestKeyFile does not exist: $KeyFile"
    }

    if ($usesUntrustedCertificate) {
        Write-Warning 'This package uses an untrusted local certificate. Windows will show an Unknown Publisher or trust warning on other computers. Do not upload the certificate file or private key to GitHub.'
    }

    return [pscustomobject]@{
        KeyFile = $KeyFile
        Thumbprint = $Thumbprint
        UsesUntrustedCertificate = $usesUntrustedCertificate
    }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$projectPath = Join-Path $repoRoot 'src\UzbekOrfoAddIn\UzbekOrfoAddIn.csproj'
$preflight = Join-Path $repoRoot 'eng\preflight.ps1'
$uiTest = Join-Path $repoRoot 'eng\test-ui.ps1'
$wordTest = Join-Path $repoRoot 'eng\test-word.ps1'
$unitTest = Join-Path $repoRoot 'eng\test-unit.ps1'

if ($Configuration -ne 'Release') {
    throw 'A production release must use the Release configuration.'
}
if ($Channel -eq 'Web' -and ($UpdateUrl -notmatch '^https://')) {
    throw 'Web updates require -UpdateUrl with an HTTPS URL.'
}
if (-not $AllowDirtyWorktree) {
    & git -C $repoRoot diff --quiet
    if ($LASTEXITCODE -ne 0) { throw 'Worktree has tracked changes. Commit or stash them before creating a release.' }
    $untracked = & git -C $repoRoot ls-files --others --exclude-standard
    if ($untracked) { throw 'Worktree has untracked files. Commit, ignore, or remove them before creating a release.' }
}
if (-not $AllowNonReleaseBranch) {
    $branch = (& git -C $repoRoot branch --show-current).Trim()
    if ($branch -notmatch '^release/.+') {
        throw "Release must run from a release/* branch; current branch is '$branch'."
    }
}

$signing = Get-SigningConfiguration `
    -RepositoryRoot $repoRoot `
    -KeyFile $ManifestKeyFile `
    -Thumbprint $ManifestCertificateThumbprint `
    -AllowUntrustedCertificate $AllowUntrustedCertificate
& $preflight
Assert-LastExitCode 'Preflight checks'

if (-not $SkipUiTests) {
    & $uiTest -Configuration $Configuration
    Assert-LastExitCode 'UI smoke tests'
}
if (-not $SkipUnitTests) {
    & $unitTest -Configuration $Configuration
    Assert-LastExitCode 'Unit tests'
}

$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path -LiteralPath $vswhere)) { throw 'Visual Studio Build Tools were not found.' }
$vsPath = (& $vswhere -latest -products * -property installationPath).Trim()
if (-not $vsPath) { throw 'No Visual Studio installation was found.' }
$msbuild = Join-Path $vsPath 'MSBuild\Current\Bin\MSBuild.exe'
if (-not (Test-Path -LiteralPath $msbuild)) { throw "MSBuild was not found: $msbuild" }

$publishArgs = @($projectPath, '/t:Publish', "/p:Configuration=$Configuration", "/p:ClickOnceUpdateChannel=$Channel")
$publishArgs += "/p:ManifestKeyFile=$($signing.KeyFile)"
$publishArgs += "/p:ManifestCertificateThumbprint=$($signing.Thumbprint)"
if ($Channel -eq 'Web') { $publishArgs += "/p:ClickOnceUpdateUrl=$UpdateUrl" }
if ($PublishDirectory) { $publishArgs += "/p:PublishDir=$PublishDirectory" }

if (-not $SkipPublish) {
    & $msbuild @publishArgs
    Assert-LastExitCode 'ClickOnce publish'
}
if (-not $SkipWordTests) {
    & $wordTest -Configuration $Configuration
    Assert-LastExitCode 'Word integration smoke test'
}

Write-Host "[PASS] Release verification completed: configuration=$Configuration; channel=$Channel"
