$ErrorActionPreference = 'Stop'

# Keep this script strictly ASCII for Windows PowerShell 5.1.
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$path = 'GenerateUserControl.cs'
$fullPath = (Resolve-Path $path).Path
$text = [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8)
$nl = if ($text.Contains("`r`n")) { "`r`n" } else { "`n" }

$beautifyAdd = '            _quickActionsPanel.Controls.Add(_beautifyButton);'
$ragAdd = '            _quickActionsPanel.Controls.Add(_beautifyUseRagCheckBox);'

if ($text.Contains($beautifyAdd) -and -not $text.Contains($ragAdd)) {
    $replacement =
        $beautifyAdd + $nl +
        '            InitializeBeautifyRagOption();' + $nl +
        $ragAdd
    $text = $text.Replace($beautifyAdd, $replacement)
} elseif (-not $text.Contains($ragAdd)) {
    throw 'Could not locate Beautify button insertion point for RAG option.'
}

[System.IO.File]::WriteAllText($fullPath, $text, $utf8NoBom)

$beautifyPath = 'GenerateUserControl.Beautify.cs'
$beautifyFullPath = (Resolve-Path $beautifyPath).Path
$beautify = [System.IO.File]::ReadAllText($beautifyFullPath, [System.Text.Encoding]::UTF8)
$oldCall = '            await BeautifySelectedWordTextAsync(preset);'
$newCall = '            await BeautifySelectedWordTextWithOptionalRagAsync(preset);'
if ($beautify.Contains($oldCall)) {
    $beautify = $beautify.Replace($oldCall, $newCall)
} elseif (-not $beautify.Contains($newCall)) {
    throw 'Could not route Beautify through optional RAG mode.'
}
[System.IO.File]::WriteAllText($beautifyFullPath, $beautify, $utf8NoBom)

$verify = [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8)
if (-not $verify.Contains('InitializeBeautifyRagOption();')) {
    throw 'Beautify RAG option initialization is missing.'
}
if (-not $verify.Contains('_quickActionsPanel.Controls.Add(_beautifyUseRagCheckBox);')) {
    throw 'Beautify RAG checkbox is not in the quick-action panel.'
}
$verifyBeautify = [System.IO.File]::ReadAllText($beautifyFullPath, [System.Text.Encoding]::UTF8)
if (-not $verifyBeautify.Contains('BeautifySelectedWordTextWithOptionalRagAsync(preset)')) {
    throw 'Beautify optional RAG routing is missing.'
}

Write-Host 'Beautify RAG option wired successfully.'
