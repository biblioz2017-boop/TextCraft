$ErrorActionPreference = 'Stop'

# Keep this script strictly ASCII for Windows PowerShell 5.1.
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$path = 'GenerateUserControl.Beautify.cs'
$fullPath = (Resolve-Path $path).Path
$text = [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8)

$oldCall = '            await BeautifyLastResponseAsync(preset);'
$newCall = '            await BeautifySelectedWordTextAsync(preset);'

if ($text.Contains($oldCall)) {
    $text = $text.Replace($oldCall, $newCall)
} elseif (-not $text.Contains($newCall)) {
    throw 'Could not switch Beautify to selected Word text mode.'
}

[System.IO.File]::WriteAllText($fullPath, $text, $utf8NoBom)

$verify = [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8)
if (-not $verify.Contains($newCall.Trim())) {
    throw 'Beautify selected Word text handler is missing.'
}
if ($verify.Contains($oldCall.Trim())) {
    throw 'Legacy last-response Beautify handler is still active.'
}

Write-Host 'Beautify now analyzes and formats selected Word text.'
