param([string]$Configuration = 'Debug', [switch]$Dark)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
$predictionVs = (& $vswhere -latest -products * -property installationPath).Trim()
$compiler = Join-Path $predictionVs 'MSBuild/Current/Bin/Roslyn/csc.exe'
$product = Join-Path $repoRoot "src/UzbekOrfoAddIn/bin/$Configuration"
$output = Join-Path $repoRoot "tests/PredictionUi/bin/$Configuration"
New-Item -ItemType Directory -Path $output -Force | Out-Null
Get-ChildItem -LiteralPath $product -Filter *.dll | Copy-Item -Destination $output -Force
$executable = Join-Path $output 'PredictionUi.exe'
& $compiler /nologo /langversion:7.3 /target:exe /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll "/r:$product/UzbekOrfoAddIn.dll" "/out:$executable" (Join-Path $repoRoot 'tests/PredictionUi/Program.cs')
if ($LASTEXITCODE -ne 0) { throw 'Prediction UI harness build failed.' }
if ($Dark) { & $executable --dark } else { & $executable }
if ($LASTEXITCODE -ne 0) { throw 'Prediction UI tests failed.' }
