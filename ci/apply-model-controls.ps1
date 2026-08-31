$ErrorActionPreference = 'Stop'

# Keep this build-time patch ASCII-only for Windows PowerShell 5.1.
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Read-Utf8Text([string]$Path) {
    $fullPath = (Resolve-Path $Path).Path
    return [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8)
}

function Write-Utf8Text([string]$Path, [string]$Text) {
    $fullPath = (Resolve-Path $Path).Path
    [System.IO.File]::WriteAllText($fullPath, $Text, $utf8NoBom)
}

Write-Host 'Applying persistent stop and Ollama unload model controls...'

$designerPath = 'Forge.Designer.cs'
$designer = Read-Utf8Text $designerPath
if (-not $designer.Contains('InitializeModelControlButtons();')) {
    $pattern = '(?m)^(\s*)InitializeComponent\(\);\s*$'
    $match = [regex]::Match($designer, $pattern)
    if (-not $match.Success) {
        throw 'Could not locate Forge constructor InitializeComponent call.'
    }

    $indent = $match.Groups[1].Value
    $replacement = $match.Value.TrimEnd() + "`r`n" + $indent + 'InitializeModelControlButtons();'
    $designer = [regex]::Replace($designer, $pattern, [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $replacement }, 1)
}
Write-Utf8Text $designerPath $designer

$forgePath = 'Forge.cs'
$forge = Read-Utf8Text $forgePath
if (-not $forge.Contains('_stopModelButton.Enabled = option;')) {
    $old = '                _optionsBox.Visible = option;'
    $new = "                _optionsBox.Visible = true;`r`n`r`n            if (_stopModelButton != null)`r`n                _stopModelButton.Enabled = option;"
    if (-not $forge.Contains($old)) {
        throw 'Could not locate CancelButtonVisibility group visibility line.'
    }
    $forge = $forge.Replace($old, $new)
}
Write-Utf8Text $forgePath $forge

Write-Host 'Model controls patch applied.'
