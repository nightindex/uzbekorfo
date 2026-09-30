param(
    [string]$CorpusPath = (Join-Path $PSScriptRoot '..\tests\Corpora\spelling_engineering_seed.json'),
    [string]$Configuration = 'Debug',
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '..\TestResults\spelling')
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$product = Join-Path $root "src\UzbekOrfoAddIn\bin\$Configuration"
if (-not (Test-Path -LiteralPath "$product\UzbekOrfoAddIn.dll")) { throw 'Build the add-in before measuring.' }
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$vsPath = & $vswhere -latest -products * -property installationPath
if (-not $vsPath) { throw 'Visual Studio C# compiler is required.' }
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
Get-ChildItem -LiteralPath $product -Filter *.dll | Copy-Item -Destination $OutputDirectory -Force
$exe = Join-Path $OutputDirectory 'SpellingEvaluation.exe'
& (Join-Path $vsPath 'MSBuild\Current\Bin\Roslyn\csc.exe') /nologo /langversion:7.3 /target:exe /r:System.Web.Extensions.dll "/r:$product\UzbekOrfoAddIn.dll" "/out:$exe" (Join-Path $root 'tests\SpellingEvaluation\Program.cs')
if ($LASTEXITCODE -ne 0) { throw 'Evaluation harness compilation failed.' }
& $exe (Join-Path $root 'src\UzbekOrfoAddIn\Data') $CorpusPath (Join-Path $OutputDirectory 'results.json')
if ($LASTEXITCODE -ne 0) { throw 'Spelling evaluation failed.' }
