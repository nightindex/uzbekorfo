param([string]$Configuration = 'Release')
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$wordTestVs = & $vswhere -latest -products * -property installationPath
if (-not $wordTestVs) { throw 'Visual Studio with the C# compiler is required.' }
$compiler = Join-Path $wordTestVs 'MSBuild\Current\Bin\Roslyn\csc.exe'
$product = Join-Path $repoRoot "src\UzbekOrfoAddIn\bin\$Configuration"
$interop = Join-Path $repoRoot "src\UzbekOrfoAddIn\obj\$Configuration\Interop.Microsoft.Office.Interop.Word.dll"
if (-not (Test-Path -LiteralPath $interop)) { throw 'Build the add-in first to generate Word interop references.' }
$outputDir = Join-Path $repoRoot "tests\WordSmoke\bin\$Configuration"
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
Get-ChildItem -LiteralPath $product -Filter *.dll | Copy-Item -Destination $outputDir
$outputExe = Join-Path $outputDir 'WordSmoke.exe'
& $compiler /nologo /langversion:7.3 /target:exe "/link:$interop" "/r:$product\UzbekOrfoAddIn.dll" "/out:$outputExe" (Join-Path $repoRoot 'tests\WordSmoke\Program.cs')
if ($LASTEXITCODE -ne 0) { throw 'Word smoke harness compilation failed.' }
# Creates a separate hidden Word instance and blank documents, closed without saving.
& $outputExe
if ($LASTEXITCODE -ne 0) { throw 'Word smoke checks failed.' }
