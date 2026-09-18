param([string]$Configuration = 'Release', [ValidateSet('cyrillic', 'latin')][string]$Script = 'cyrillic', [switch]$PilotEvaluation, [switch]$IncludeCourt)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
$corpusVs = (& $vswhere -latest -products * -property installationPath).Trim()
$corpusCompiler = Join-Path $corpusVs 'MSBuild/Current/Bin/Roslyn'
$project = Join-Path $repoRoot 'tools/MatnAi.Corpus/MatnAi.Corpus.csproj'
& dotnet build $project -c $Configuration /p:UseSharedCompilation=false "/p:CscToolPath=$corpusCompiler" /p:CscToolExe=csc.exe /clp:ErrorsOnly
if ($LASTEXITCODE -ne 0) { throw 'Corpus tool build failed.' }
$corpusArguments = @($repoRoot, $Script)
if ($PilotEvaluation) { $corpusArguments += '--pilot-evaluation' }
if ($IncludeCourt) { $corpusArguments += '--include-court' }
& dotnet (Join-Path $repoRoot "tools/MatnAi.Corpus/bin/$Configuration/net8.0/MatnAi.Corpus.dll") @corpusArguments
if ($LASTEXITCODE -ne 0) { throw "Corpus preparation or extraction failed ($Script); inspect docs/prediction reports." }
