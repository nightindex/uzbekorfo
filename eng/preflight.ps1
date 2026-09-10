param(
    [string]$ProjectPath = (Join-Path $PSScriptRoot "..\\src\\UzbekOrfoAddIn\\UzbekOrfoAddIn.csproj")
)

$ErrorActionPreference = "Stop"

function Write-Check([string]$status, [string]$message) {
    Write-Host "[$status] $message"
}

if (-not (Test-Path $ProjectPath)) {
    Write-Check "FAIL" "Project file not found: $ProjectPath"
    exit 1
}

$failures = New-Object System.Collections.Generic.List[string]
$warnings = New-Object System.Collections.Generic.List[string]

$projectDir = Split-Path -Parent (Resolve-Path $ProjectPath)
$repoRoot = Split-Path -Parent (Split-Path -Parent $projectDir)
if (-not (Test-Path (Join-Path $repoRoot ".gitignore"))) {
    $repoRoot = $projectDir
}
$dataDir = Join-Path $projectDir "Data"

$requiredData = @(
    "uzbek_main.dic",
    "translit_exceptions.json",
    "explanations.json",
    "uzbek_suffixes.json",
    "grammar_rules.json",
    "proper_nouns.json",
    "uzbek_dictionary.json"
)

foreach ($f in $requiredData) {
    $path = Join-Path $dataDir $f
    if (-not (Test-Path $path)) {
        $failures.Add("Missing required data file: $path")
    }
}

$projXml = [xml](Get-Content -Raw $ProjectPath)
$projText = Get-Content -Raw $ProjectPath

if ($projText -match "(?i)<ManifestKeyFile>[^<]*temp[^<]*\\.pfx</ManifestKeyFile>") {
    $failures.Add("Project still references a temporary .pfx manifest key in the project file.")
}

if ($projText -match "(?i)<(?:None|Content|EmbeddedResource)\s+Include=""[^""]*\.pfx""") {
    $failures.Add("Project file includes a .pfx certificate. Keep signing keys outside source control.")
}

$hasUnconditionalSignTrue = $projText -match "<PropertyGroup>\s*<SignManifests>\s*true\s*</SignManifests>\s*</PropertyGroup>"
if ($hasUnconditionalSignTrue -and -not ($projText -match "<ManifestKeyFile>.+</ManifestKeyFile>")) {
    $failures.Add("Unconditional SignManifests=true detected without ManifestKeyFile.")
}

$gitignorePath = Join-Path $repoRoot ".gitignore"
if (-not (Test-Path $gitignorePath)) {
    $failures.Add("Missing .gitignore. Build outputs and signing keys must be ignored before pushing.")
} else {
    $gitignoreText = Get-Content -Raw $gitignorePath
    foreach ($pattern in @("*.pfx", "bin/", "obj/", ".vs/", "publish/", "Directory.Build.props")) {
        if ($gitignoreText -notmatch [regex]::Escape($pattern)) {
            $failures.Add(".gitignore does not protect required pattern: $pattern")
        }
    }
}

$vsToolsSearchRoots = @(
    (Join-Path ${env:ProgramFiles} "Microsoft Visual Studio"),
    (Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio"),
    (Join-Path ${env:ProgramFiles(x86)} "MSBuild\\Microsoft\\VisualStudio")
) | Where-Object { $_ -and (Test-Path $_) }

$vsToolsTarget = $null
foreach ($root in $vsToolsSearchRoots) {
    $match = Get-ChildItem -Path $root -Recurse -Filter "Microsoft.VisualStudio.Tools.Office.targets" -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty FullName
    if ($match) {
        $vsToolsTarget = $match
        break
    }
}

if (-not $vsToolsTarget) {
    $warnings.Add("VSTO Office targets were not found under standard Visual Studio/MSBuild installation roots.")
}

$localProps = Join-Path $repoRoot "Directory.Build.props"
if (Test-Path $localProps) {
    $propsText = Get-Content -Raw $localProps
    if ($propsText -match "REPLACE_WITH_REAL_THUMBPRINT") {
        $warnings.Add("Directory.Build.props still contains placeholder certificate thumbprint.")
    }
} else {
    $warnings.Add("Directory.Build.props not found. This is expected for clean clones; copy Directory.Build.props.example locally when signing a release.")
}

$testsDir = Join-Path $repoRoot "tests"
$testFiles = Get-ChildItem -Path $testsDir -Recurse -File -Include *.cs,*.csproj -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match "Test|Tests" } |
    Select-Object -First 1
if (-not $testFiles) {
    $warnings.Add("No obvious automated test files/projects detected.")
}

if ($failures.Count -eq 0) {
    Write-Check "PASS" "Preflight mandatory checks passed."
} else {
    foreach ($f in $failures) { Write-Check "FAIL" $f }
}

foreach ($w in $warnings) { Write-Check "WARN" $w }

if ($failures.Count -gt 0) { exit 1 }
exit 0
