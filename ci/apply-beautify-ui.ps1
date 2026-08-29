$ErrorActionPreference = 'Stop'

# ASCII-only build patch for Windows PowerShell 5.1.
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$path = 'GenerateUserControl.cs'
$fullPath = (Resolve-Path $path).Path
$text = [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8)
$nl = if ($text.Contains("`r`n")) { "`r`n" } else { "`n" }

$matrixAdd = '            _quickActionsPanel.Controls.Add(_matrixButton);'
$beautifyAdd = '            _quickActionsPanel.Controls.Add(_beautifyButton);'

if ($text.Contains($matrixAdd)) {
    $replacement =
        '            InitializeBeautifyButton();' + $nl +
        '            _quickActionsPanel.Controls.Add(_beautifyButton);'
    $text = $text.Replace($matrixAdd, $replacement)
} elseif (-not $text.Contains($beautifyAdd)) {
    throw 'Could not locate quick-action insertion point for Beautify.'
}

[System.IO.File]::WriteAllText($fullPath, $text, $utf8NoBom)

$verify = [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8)
if (-not $verify.Contains('InitializeBeautifyButton();')) {
    throw 'Beautify button initialization is missing.'
}
if (-not $verify.Contains('_quickActionsPanel.Controls.Add(_beautifyButton);')) {
    throw 'Beautify button is not in the quick-action panel.'
}
if ($verify.Contains('_quickActionsPanel.Controls.Add(_matrixButton);')) {
    throw 'Standalone Matrix button is still visible in the quick-action panel.'
}

Write-Host 'Standalone Matrix button removed; Beautify presets button enabled.'
