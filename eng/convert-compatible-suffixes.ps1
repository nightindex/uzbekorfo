param(
    [string]$SourceDirectory = (Join-Path $PSScriptRoot '..\obj\restored-hunspell'),
    [string]$DataDirectory = (Join-Path $PSScriptRoot '..\src\UzbekOrfoAddIn\Data'),
    [string]$AssemblyPath = (Join-Path $PSScriptRoot '..\src\UzbekOrfoAddIn\bin\Debug\UzbekOrfoAddIn.dll')
)
$ErrorActionPreference = 'Stop'
[void][Reflection.Assembly]::LoadFrom([IO.Path]::GetFullPath($AssemblyPath))
$transliterator = [UzbekOrfoAddIn.Services.TransliterationService]::new('')
[void][Reflection.Assembly]::LoadWithPartialName('System.Web.Extensions')
$serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$serializer.MaxJsonLength = [int]::MaxValue
$utf8 = New-Object Text.UTF8Encoding($false, $true)
$path = Join-Path $DataDirectory 'uzbek_suffixes.json'
$payload = $serializer.DeserializeObject([IO.File]::ReadAllText($path, $utf8))
$latin = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
function Normalize-Apostrophes([string]$value) {
    return $value.Replace([char]0x02bb, "'").Replace([char]0x02bc, "'").Replace([char]0x2018, "'").Replace([char]0x2019, "'")
}
foreach ($line in [IO.File]::ReadLines((Join-Path $SourceDirectory 'suf_uz_UZ.txt'), $utf8)) {
    if ($line -cmatch '^SFX (\S) 0 ([^/ ]+) \.$') {
        [void]$latin.Add($Matches[1] + '|' + (Normalize-Apostrophes $Matches[2]))
    }
}
$baseSuffixes = @($payload.suffixes | Where-Object { $_.id -notlike 'IMPORTED_*' })
$existing = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
foreach ($suffix in $baseSuffixes) { [void]$existing.Add($suffix.cyrillic) }
$mapped = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::Ordinal)
$counts = [ordered]@{ SourceRules=0; NeedsStrippingOrConditionOrContinuation=0; MissingLatinCounterpart=0; ExistingSurface=0; MappedSourceRows=0 }
foreach ($line in [IO.File]::ReadLines((Join-Path $SourceDirectory 'suf_uz_UZ_Cyrl.txt'), $utf8)) {
    if ($line -notmatch '^SFX \S+ \S+ \S+ \S+') { continue }
    $counts.SourceRules++
    if ($line -cnotmatch '^SFX (\S) 0 ([^/ ]+) \.$') { $counts.NeedsStrippingOrConditionOrContinuation++; continue }
    $flag = $Matches[1]; $form = $Matches[2]
    if ($form -eq '0') { $counts.NeedsStrippingOrConditionOrContinuation++; continue }
    if ($existing.Contains($form)) { $counts.ExistingSurface++; continue }
    $latinForm = Normalize-Apostrophes ($transliterator.ToLatin($form))
    if (-not $latin.Contains($flag + '|' + $latinForm)) { $counts.MissingLatinCounterpart++; continue }
    if (-not $mapped.ContainsKey($form)) {
        $mapped.Add($form, [pscustomobject]@{ Cyrillic=$form; Latin=$latinForm; Flags=(New-Object 'System.Collections.Generic.SortedSet[char]') })
    }
    [void]$mapped[$form].Flags.Add([char]$flag)
    $counts.MappedSourceRows++
}
$additions = New-Object 'System.Collections.Generic.List[object]'
$index = 0
foreach ($item in ($mapped.Values | Sort-Object Cyrillic)) {
    $index++
    $additions.Add([ordered]@{
        id=('IMPORTED_{0:D4}' -f $index); cyrillic=$item.Cyrillic; latin=$item.Latin
        category='imported_complete'; family='imported'; order=1; productive=$true
        standaloneOnly=$true; requiredRootFlags=(-join $item.Flags)
        description='Complete source suffix; requires a matching root flag; cannot be chained.'
    })
}
$payload.suffixes = @($baseSuffixes) + $additions.ToArray()
$families = @($payload.families | Where-Object { $_.id -ne 'imported' })
$payload.families = $families + @([ordered]@{
    id='imported'; rootLookupSuffix=''; allowedRootPartOfSpeech=@('noun','verb','adjective','adverb','pronoun','numeral','proper_noun','abbreviation','conjunction','postposition','particle','interjection'); allowUnknownRootPartOfSpeech=$true
    categoryStages=@{ imported_complete=1 }; occurrenceLimits=@(@{categories=@('imported_complete');max=1}); requirements=@()
})
$temp = "$path.tmp"
[IO.File]::WriteAllText($temp, (ConvertTo-Json -InputObject $payload -Depth 12) + [Environment]::NewLine, $utf8)
Move-Item -LiteralPath $temp -Destination $path -Force
$report = [ordered]@{ Counts=$counts; ImportedSuffixes=$additions.Count; OriginalSuffixes=$baseSuffixes.Count; TotalSuffixes=$payload.suffixes.Count }
[IO.File]::WriteAllText((Join-Path $DataDirectory 'suffix-migration-report.json'), (ConvertTo-Json -InputObject $report -Depth 5), $utf8)
Write-Host ($report | ConvertTo-Json -Depth 5)
