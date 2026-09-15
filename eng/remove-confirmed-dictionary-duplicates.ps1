<# One-time, exact-key data migration. Never removes other screening findings. #>
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$path = Join-Path $root 'src\UzbekOrfoAddIn\Data\uzbek_dictionary.json'
$backup = Join-Path $root 'TestResults\dictionary-before-confirmed-removals.json'
$utf8 = New-Object Text.UTF8Encoding($false, $true)
$text = [IO.File]::ReadAllText($path, $utf8)
$data = $text | ConvertFrom-Json
$pairs = @(@('атакcия','атаксия'), @('юнeско','юнеско'))
if (@($data.Entries | Where-Object { $_.Word -ceq 'атакcия' -or $_.Word -ceq 'юнeско' }).Count -eq 0) {
    Write-Host 'Already applied; no changes.'
    return
}
foreach ($pair in $pairs) {
    foreach ($key in $pair) {
        if (@($data.Entries | Where-Object { $_.Word -ceq $key }).Count -ne 1) { throw "Expected exactly one entry: $key" }
    }
}
$source = $data.Entries | Where-Object { $_.Word -ceq 'юнeско' }
$target = $data.Entries | Where-Object { $_.Word -ceq 'юнеско' }
if ($null -ne $target.Definition) { throw 'Target definition changed; review required.' }
foreach ($pair in $pairs) {
    $pattern = '(?m)^                    \{\r?\n                        "Word":  "' +
        [regex]::Escape($pair[0]) + '",[\s\S]*?^                    \},\r?\n'
    if ([regex]::Matches($text, $pattern).Count -ne 1) { throw 'Unexpected source layout.' }
    $text = [regex]::Replace($text, $pattern, '')
}
$pattern = '("Word":  "юнеско",\r?\n                        "Definition":  )null'
if ([regex]::Matches($text, $pattern).Count -ne 1) { throw 'Unexpected target layout.' }
$definitionJson = ConvertTo-Json -InputObject $source.Definition -Compress
$text = [regex]::Replace($text, $pattern, [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $m.Groups[1].Value + $definitionJson })
$text = $text.Replace('"WordCount":  216341', '"WordCount":  216339')
$text = $text.Replace('"DataVersion":  "2026.09.14.4"', '"DataVersion":  "2026.09.15.1"')
$text = $text.Replace('"UpdatedAtUtc":  "2026-09-14T11:44:00Z"', ('"UpdatedAtUtc":  "' + [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ') + '"'))
$after = $text | ConvertFrom-Json
if (@($after.Entries).Count -ne ($data.Entries.Count - 2) -or $after.WordCount -ne $after.Entries.Count) { throw 'Unexpected entry count.' }
if (@($after.Entries | Where-Object Definition).Count -ne $after.DefinitionCount) { throw 'Definition count changed unexpectedly.' }
New-Item -ItemType Directory -Path (Split-Path -Parent $backup) -Force | Out-Null
if (Test-Path -LiteralPath $backup) { throw "Recovery copy already exists: $backup" }
Copy-Item -LiteralPath $path -Destination $backup
[IO.File]::WriteAllText($path, $text, $utf8)
Write-Host "Removed two exact mixed-script duplicates; definition preserved. Recovery copy: $backup"
