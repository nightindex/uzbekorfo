# Isolated regression fixture: never modifies the bundled dictionary.
$ErrorActionPreference = 'Stop'
$fixture = Join-Path $PSScriptRoot ("..\obj\hunspell-test-" + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($fixture)
$utf8 = New-Object Text.UTF8Encoding($false, $true)
$canonical = Join-Path $fixture 'uzbek_dictionary.json'
$latin = Join-Path $fixture 'latin.dic'
$cyrillic = Join-Path $fixture 'cyrillic.dic'
$special = [char]0x044F
[IO.File]::WriteAllText($canonical, '{"Schema":"uzbekorfo-dictionary-v2","WordCount":1,"DefinitionCount":0,"Entries":[{"Word":"kitob","Lemma":"kitob","PartOfSpeech":"noun","HunspellFlags":"a"}]}', $utf8)
[IO.File]::WriteAllText($latin, "5`nkitob/A`nkitob/a`nsodda`nmaxsus/A$special`nbelgi/-`n", $utf8)
[IO.File]::WriteAllText($cyrillic, "1`nkitob/A`n", $utf8)
$arguments = @{ LatinPath=$latin; CyrillicPath=$cyrillic; DataDirectory=$fixture; ReportDirectory=$fixture }
$before = (Get-FileHash -LiteralPath $canonical).Hash
& "$PSScriptRoot\import-hunspell-dictionaries.ps1" @arguments
if ((Get-FileHash -LiteralPath $canonical).Hash -ne $before) { throw 'Dry run changed canonical data.' }
& "$PSScriptRoot\import-hunspell-dictionaries.ps1" @arguments -Apply
$payload = [IO.File]::ReadAllText($canonical, $utf8) | ConvertFrom-Json
$root = $payload.Entries | Where-Object Word -eq 'kitob'
if ($root.HunspellFlags -cne 'Aa' -or $root.PartOfSpeech -ne 'noun' -or $payload.WordCount -ne 4) { throw 'Flag merge or unflagged import failed.' }
if (($payload.Entries | Where-Object Word -eq 'maxsus').HunspellFlags -cne "A$special") { throw 'Special flag lost.' }
if (($payload.Entries | Where-Object Word -eq 'belgi').HunspellFlags -cne '-') { throw 'Dash flag lost.' }
& "$PSScriptRoot\sync-dictionary-data.ps1" -DataDirectory $fixture
& "$PSScriptRoot\sync-dictionary-data.ps1" -DataDirectory $fixture -Check
$before = (Get-FileHash -LiteralPath $canonical).Hash
& "$PSScriptRoot\import-hunspell-dictionaries.ps1" @arguments -Apply
if ((Get-FileHash -LiteralPath $canonical).Hash -ne $before) { throw 'Second import is not idempotent.' }
Write-Host "Hunspell import regression checks passed. Fixture: $fixture"
