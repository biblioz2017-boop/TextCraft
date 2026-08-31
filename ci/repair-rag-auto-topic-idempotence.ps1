$ErrorActionPreference = 'Stop'

# Keep this script strictly ASCII for Windows PowerShell 5.1.
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$path = 'ci\apply-rag-auto-topic.ps1'
$fullPath = (Resolve-Path $path).Path
$text = [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8)
$nl = if ($text.Contains("`r`n")) { "`r`n" } else { "`n" }

$branchPattern = @'
(?s)\}\s*elseif \(-not \$science\.Contains\('string sourceTopicSeed = await rag\.BuildCheckedPdfTopicSeedAsync\(\);'\)\) \{\s*throw 'Could not locate strict-RAG evidence retrieval call for auto-topic fallback\.'\s*\}
'@
$branchReplacement = @'
} elseif ($science.Contains('string sourceTopicSeed = await rag.BuildCheckedPdfTopicSeedAsync();')) {
    Write-Host 'Automatic RAG topic discovery is already active.'
} elseif ($science.Contains('GetCheckedPdfOverviewEvidence(10)')) {
    Write-Host 'Balanced overview routing is already active; auto-topic fallback is superseded.'
} else {
    throw 'Could not locate strict-RAG evidence retrieval call for auto-topic fallback.'
}
'@

if ([regex]::IsMatch($text, $branchPattern.Trim())) {
    $replacement = $branchReplacement.Trim().Replace("`n", $nl)
    $text = [regex]::Replace($text, $branchPattern.Trim(), [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $replacement }, 1)
} elseif (-not $text.Contains('Balanced overview routing is already active; auto-topic fallback is superseded.')) {
    throw 'Could not make auto-topic branch idempotent.'
}

$validationPattern = @'
(?s)if \(-not \(Read-Utf8Text \$sciencePath\)\.Contains\('sourceTopicSeed = await rag\.BuildCheckedPdfTopicSeedAsync'\)\) \{\s*throw 'Automatic topic fallback is not wired into strict RAG\.'\s*\}
'@
$validationReplacement = @'
$scienceAfter = Read-Utf8Text $sciencePath
if (-not (
    $scienceAfter.Contains('sourceTopicSeed = await rag.BuildCheckedPdfTopicSeedAsync') -or
    $scienceAfter.Contains('GetCheckedPdfOverviewEvidence(10)')
)) {
    throw 'Automatic topic or balanced overview routing is not wired into strict RAG.'
}
'@

if ([regex]::IsMatch($text, $validationPattern.Trim())) {
    $replacement = $validationReplacement.Trim().Replace("`n", $nl)
    $text = [regex]::Replace($text, $validationPattern.Trim(), [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $replacement }, 1)
} elseif (-not $text.Contains('Automatic topic or balanced overview routing is not wired into strict RAG.')) {
    throw 'Could not make auto-topic validation idempotent.'
}

[System.IO.File]::WriteAllText($fullPath, $text, $utf8NoBom)
Write-Host 'Auto-topic patch made idempotent for multi-pass MSI builds.'
