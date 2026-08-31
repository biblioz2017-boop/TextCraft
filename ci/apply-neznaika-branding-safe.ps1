$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

$owlPath = Join-Path $root 'Assets\NeZnaikaOwl.jpg'
$hiddenOwlPath = Join-Path $root 'Assets\NeZnaikaOwl.jpg.ci-hidden'
$brandingScript = Join-Path $root 'ci\apply-neznaika-branding.ps1'
$resxPath = Join-Path $root 'AboutBox.resx'

if (-not (Test-Path $owlPath)) { throw 'NeZnaika owl source image is missing.' }
if (-not (Test-Path $brandingScript)) { throw 'Branding script is missing.' }
if (-not (Test-Path $resxPath)) { throw 'AboutBox.resx is missing.' }

# Hide the JPEG only while the legacy branding script runs so that script cannot
# attempt a GDI+ decode. The original JPEG is embedded later as raw bytes.
Move-Item -LiteralPath $owlPath -Destination $hiddenOwlPath -Force
try {
    $scriptText = [System.IO.File]::ReadAllText($brandingScript, [System.Text.Encoding]::UTF8)
    Invoke-Expression $scriptText
}
finally {
    if (Test-Path $hiddenOwlPath) {
        Move-Item -LiteralPath $hiddenOwlPath -Destination $owlPath -Force
    }
}

# Validate only the JPEG signature. Do not decode or re-encode the image in CI.
$owlBytes = [System.IO.File]::ReadAllBytes($owlPath)
if ($owlBytes.Length -lt 4 -or $owlBytes[0] -ne 0xFF -or $owlBytes[1] -ne 0xD8 -or $owlBytes[$owlBytes.Length - 2] -ne 0xFF -or $owlBytes[$owlBytes.Length - 1] -ne 0xD9) {
    throw 'NeZnaika owl source does not have a valid JPEG signature.'
}

# Remove the legacy WinForms Image object from the resx. The original JPEG is a raw
# EmbeddedResource and therefore never passes through GenerateResource/GDI+.
$resx = [System.IO.File]::ReadAllText($resxPath, [System.Text.Encoding]::UTF8)
$logoPattern = '(?s)\s*<data name="logoPictureBox.Image"[^>]*>.*?</data>\s*'
if ([regex]::IsMatch($resx, $logoPattern)) {
    $resx = [regex]::Replace($resx, $logoPattern, "`r`n", 1)
}
if ($resx.Contains('name="logoPictureBox.Image"')) {
    throw 'Failed to remove logoPictureBox.Image from AboutBox.resx.'
}

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($resxPath, $resx, $utf8NoBom)

Write-Host ('NeZnaika branding applied; original JPEG validated and kept as raw manifest resource (' + $owlBytes.Length + ' bytes).')
