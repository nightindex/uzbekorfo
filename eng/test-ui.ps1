param([string]$Configuration = "Release")
$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
& (Join-Path $PSScriptRoot 'test-ribbon-layout.ps1')
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$vsPath = & $vswhere -latest -products * -property installationPath
if (-not $vsPath) { throw 'Visual Studio with the C# compiler is required.' }
$compiler = Join-Path $vsPath 'MSBuild\Current\Bin\Roslyn\csc.exe'
$outputDir = Join-Path $repoRoot "tests\UiCompatibility\bin\$Configuration"
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
$sources = @(
    'UI\ScreenGeometry.cs', 'UI\DpiLayout.cs', 'UI\OverlayPositioner.cs', 'UI\DpiForm.cs',
    'UI\ThemeManager.cs', 'UI\ModernForm.cs',
    'UI\Controls\ModernButton.cs', 'UI\Controls\ModernCard.cs',
    'UI\Controls\EmojiLabel.cs', 'UI\Controls\DpiDataGridView.cs', 'UI\Controls\ModernScrollBar.cs',
    'UI\Controls\ModernTextBox.cs', 'UI\Controls\ModernToggle.cs',
    'UI\Controls\ToastNotification.cs', 'UI\ModernMessageBox.cs',
    'Helpers\AnimationHelper.cs', 'Forms\AddNewWordsForm.cs',
    'Forms\AppInfoForm.cs', 'Services\AppUpdateService.cs', 'Forms\ImportResultForm.cs', 'Forms\ImportProgressForm.cs',
    'Forms\MatnAiSettingsForm.cs', 'Forms\MatnAiHelpForm.cs',
    'Forms\TranslitExceptionsForm.cs', 'Forms\ViewErrorsForm.cs',
    'Models\TranslitException.cs', 'Models\Enums.cs', 'Models\ErrorEntry.cs', 'Models\Suggestion.cs',
    'Core\ITransliterator.cs', 'Services\TransliterationService.cs',
    'Helpers\TextHelper.cs', 'Helpers\Logger.cs', 'Helpers\HotkeyManager.cs'
    'Forms\CompletionPopup.cs', 'Forms\GhostSuggestionWindow.cs'
) | ForEach-Object { Join-Path $repoRoot "src\UzbekOrfoAddIn\$_" }
$sources += Join-Path $repoRoot 'tests\UiCompatibility\Program.cs'
$sources += Join-Path $repoRoot 'tests\UiCompatibility\HostGuards.cs'
$outputExe = Join-Path $outputDir 'UiCompatibility.exe'
$jsonAssembly = Join-Path $repoRoot "src\UzbekOrfoAddIn\bin\$Configuration\Newtonsoft.Json.dll"
if (-not (Test-Path -LiteralPath $jsonAssembly)) {
    $packageRoot = if ($env:NUGET_PACKAGES) { $env:NUGET_PACKAGES } else { Join-Path $env:USERPROFILE '.nuget/packages' }
    $jsonAssembly = Join-Path $packageRoot 'newtonsoft.json/13.0.3/lib/net45/Newtonsoft.Json.dll'
}
if (-not (Test-Path -LiteralPath $jsonAssembly)) { throw 'Restore test dependencies before running UI tests (Newtonsoft.Json required).' }
Copy-Item -LiteralPath $jsonAssembly -Destination $outputDir -Force
& $compiler /nologo /langversion:7.3 /target:exe /r:Accessibility.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.Deployment.dll /r:System.Net.Http.dll "/r:$jsonAssembly" /r:System.Xml.Linq.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll /r:Microsoft.CSharp.dll "/out:$outputExe" $sources
if ($LASTEXITCODE -ne 0) { throw 'UI harness compilation failed.' }
& $outputExe
if ($LASTEXITCODE -ne 0) { throw 'UI compatibility checks failed.' }
