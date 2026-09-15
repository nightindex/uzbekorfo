<# Static guard for the layout consumed by the VSTO Ribbon Designer, not an IDE automation test. #>
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$designer = Get-Content -LiteralPath (Join-Path $root 'src\UzbekOrfoAddIn\UzbekOrfoRibbon.Designer.cs') -Raw -Encoding UTF8
$handlers = Get-Content -LiteralPath (Join-Path $root 'src\UzbekOrfoAddIn\UzbekOrfoRibbon.MatnAi.cs') -Raw -Encoding UTF8
if ($designer.Contains('InitializeMatnAi(') -or $handlers.Contains('CreateRibbon')) {
    throw 'MatnAi controls must be declared in InitializeComponent, not a runtime-only initializer.'
}
if ($designer.Contains('btnMatnAiShow') -or $designer.Contains('matnaiLearning') -or
    $designer.Contains('btnMatnAiReset') -or $designer.Contains('btnMatnAiRebuild')) {
    throw 'Only the primary toggle, settings and help controls belong in the MatnAI ribbon group.'
}
if (-not $handlers.Contains('new MatnAiSettingsForm(') -or -not $handlers.Contains('new MatnAiHelpForm()') -or
    -not $handlers.Contains('MatnAiLearningConsent') -or -not $handlers.Contains('ResetLearning()') -or
    -not $handlers.Contains('Rebuild()')) {
    throw 'MatnAI settings must retain personal learning, reset and manual index refresh actions.'
}
if ([regex]::Matches($designer, 'this\.Tabs\.Add\(').Count -ne 1 -or
    -not $designer.Contains('this.Tabs.Add(this.tabUzbekOrfo);') -or $designer.Contains('tabMatnAi')) {
    throw 'MatnAi must use the existing Uzbek Orfo tab, not a separate tab.'
}
if ([regex]::Matches($designer, [regex]::Escape('this.tabUzbekOrfo.Groups.Add(this.groupMatnAi);')).Count -ne 1) {
    throw 'Expected one MatnAi group in the Uzbek Orfo tab.'
}
if (-not $designer.Contains('this.groupMatnAi.Label = "MatnAI";')) {
    throw 'The feature group must use the MatnAI product name.'
}
foreach ($control in @('matnaiEnabled', 'btnMatnAiSettings', 'btnMatnAiHelp')) {
    if (-not $designer.Contains("this.$control = this.Factory.CreateRibbon") -or
        [regex]::Matches($designer, [regex]::Escape("this.groupMatnAi.Items.Add(this.$control);")).Count -ne 1 -or
        -not $designer.Contains("RibbonControlEventHandler(this.$($control)_Click)") -or
        -not $handlers.Contains("void $($control)_Click(")) { throw "Incomplete designer control/handler: $control" }
    if (-not $designer.Contains("this.$control.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;") -or
        -not $designer.Contains("this.$control.ShowImage = true;") -or
        -not [regex]::IsMatch($designer, ('this\.' + [regex]::Escape($control) + '\.OfficeImageId = "[^"]+";'))) {
        throw "MatnAi control must have a large icon and visible label: $control"
    }
    $labelPattern = 'this\.' + [regex]::Escape($control) + '\.Label = "[^\"]*\p{IsCyrillic}[^\"]*";'
    $screenTipPattern = 'this\.' + [regex]::Escape($control) + '\.ScreenTip = "[^\"]*\p{IsCyrillic}[^\"]*";'
    # A Visual Studio Designer may split a long SuperTip over multiple C# string literals.
    $superTipPattern = 'this\.' + [regex]::Escape($control) + '\.SuperTip = "[^\"]*\p{IsCyrillic}'
    if (-not [regex]::IsMatch($designer, $labelPattern) -or
        -not [regex]::IsMatch($designer, $screenTipPattern) -or
        -not [regex]::IsMatch($designer, $superTipPattern)) {
        throw "MatnAI control must have a Cyrillic Uzbek label, ScreenTip and SuperTip: $control"
    }
}
Write-Host 'PASS: MatnAI has three large, fully described Uzbek Cyrillic controls in the Uzbek Orfo tab.'
