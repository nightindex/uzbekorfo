param(
    [string]$Configuration = 'Release',
    [string]$ProjectPath
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $ProjectPath) {
    $ProjectPath = Join-Path $repoRoot 'tests\UzbekOrfoAddIn.UnitTests\UzbekOrfoAddIn.UnitTests.csproj'
}
if (-not (Test-Path -LiteralPath $ProjectPath)) {
    throw "Unit-test project was not found: $ProjectPath"
}

$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path -LiteralPath $vswhere)) {
    throw 'Visual Studio is required to run unit tests with the supported compiler.'
}
$vsPath = (& $vswhere -latest -products * -property installationPath).Trim()
$compilerDirectory = Join-Path $vsPath 'MSBuild\Current\Bin\Roslyn'
if (-not $vsPath -or -not (Test-Path -LiteralPath (Join-Path $compilerDirectory 'csc.exe'))) {
    throw 'Visual Studio C# compiler was not found.'
}

& dotnet test $ProjectPath --configuration $Configuration /p:UseSharedCompilation=false "/p:CscToolPath=$compilerDirectory" /p:CscToolExe=csc.exe
if ($LASTEXITCODE -ne 0) {
    throw "Unit tests failed with exit code $LASTEXITCODE."
}
