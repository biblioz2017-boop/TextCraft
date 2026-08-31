$ErrorActionPreference = 'Stop'

# Keep this script strictly ASCII for Windows PowerShell 5.1.
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$path = 'GenerateUserControl.Beautify.cs'
$fullPath = (Resolve-Path $path).Path
$text = [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8)

$oldCall = '            await BeautifyLastResponseAsync(preset);'
$selectionCall = '            await BeautifySelectedWordTextAsync(preset);'
$ragCall = '            await BeautifySelectedWordTextWithOptionalRagAsync(preset);'

if ($text.Contains($oldCall)) {
    $text = $text.Replace($oldCall, $selectionCall)
    Write-Host 'Beautify switched from last response to selected Word text.'
} elseif ($text.Contains($selectionCall)) {
    Write-Host 'Beautify selected Word text mode is already active.'
} elseif ($text.Contains($ragCall)) {
    Write-Host 'Beautify selected Word text mode is already superseded by optional RAG mode.'
} else {
    throw 'Could not recognize the active Beautify handler.'
}

[System.IO.File]::WriteAllText($fullPath, $text, $utf8NoBom)

$verify = [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8)
if ($verify.Contains($oldCall.Trim())) {
    throw 'Legacy last-response Beautify handler is still active.'
}
if (-not ($verify.Contains($selectionCall.Trim()) -or $verify.Contains($ragCall.Trim()))) {
    throw 'Neither selected Word text nor optional RAG Beautify handler is active.'
}

Write-Host 'Beautify selection switch is idempotent for multi-pass MSI builds.'
