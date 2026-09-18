param([string]$MetricsPath = (Join-Path $env:LOCALAPPDATA 'UzbekOrfo\MatnAI\metrics-v1.tsv'))
$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath $MetricsPath)) {
    throw 'No local metrics snapshot. Enable local statistics in MatnAI settings, use suggestions, then close Word normally.'
}
$lines = [IO.File]::ReadAllLines($MetricsPath)
if ($lines.Length -eq 0 -or $lines[0] -ne 'matnai-metrics-v1-current-session') {
    throw 'Unsupported metrics format.'
}
$counts = [ordered]@{}
$allowed = @('queries', 'shown', 'accepted', 'dismissed', 'typed_past_or_caret_moved',
    'no_candidate', 'not_displayed', 'alternatives_opened', 'insertion_rejected',
    'immediate_keyboard_undo', 'latency_le_50ms', 'latency_le_100ms',
    'latency_le_150ms', 'latency_gt_150ms')
foreach ($name in $allowed) { $counts[$name] = [long]0 }
$seen = @{}
foreach ($line in ($lines | Select-Object -Skip 1)) {
    $parts = $line.Split("`t")
    $number = [long]0
    if ($parts.Length -ne 2 -or $allowed -notcontains $parts[0] -or
        -not [long]::TryParse($parts[1], [ref]$number) -or $number -lt 0 -or $seen.ContainsKey($parts[0])) {
        throw 'Invalid or unknown metric row; report only recognized aggregate counters.'
    }
    $seen[$parts[0]] = $true
    $counts[$parts[0]] = $number
}
function Ratio([long]$numerator, [long]$denominator) {
    if ($denominator -eq 0) { return $null }
    return [Math]::Round($numerator / [double]$denominator, 4)
}
[ordered]@{
    schema = 'matnai-local-pilot-report-v1'
    scope = 'Latest saved Word session; counters are opt-in. No document text or paths included.'
    counts = $counts
    acceptancePerShown = (Ratio $counts.accepted $counts.shown)
    explicitDismissalPerShown = (Ratio $counts.dismissed $counts.shown)
    immediateKeyboardUndoPerAcceptance = (Ratio $counts.immediate_keyboard_undo $counts.accepted)
    limitations = @('Not a longitudinal study; export each session before another overwrites the snapshot.',
        'Typing past and caret movement are combined, not an explicit rejection.',
        'Undo counts only detected Ctrl+Z within five seconds, not every Undo path.',
        'Latency starts at detected context change, excluding polling delay; not full typing-to-ghost latency.',
        'Counters do not measure time saved or legal correctness.')
} | ConvertTo-Json -Depth 4
