$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

$owlPath = Join-Path $root 'Assets\NeZnaikaOwl.jpg'
$normalizedOwlPath = Join-Path $root 'Assets\NeZnaikaOwl.normalized.png'
$hiddenOwlPath = Join-Path $root 'Assets\NeZnaikaOwl.jpg.ci-hidden'
$brandingScript = Join-Path $root 'ci\apply-neznaika-branding.ps1'
$resxPath = Join-Path $root 'AboutBox.resx'

if (-not (Test-Path $owlPath)) { throw 'NeZnaika owl source image is missing.' }
if (-not (Test-Path $normalizedOwlPath)) { throw 'Normalized NeZnaika owl PNG resource is missing.' }
if (-not (Test-Path $brandingScript)) { throw 'Branding script is missing.' }
if (-not (Test-Path $resxPath)) { throw 'AboutBox.resx is missing.' }

# The legacy branding script also tries to decode the source JPEG with GDI+.
# Windows PowerShell 5.1 / x86 has proven unreliable for that operation in CI.
# Temporarily hide the JPEG while the legacy text/deployment branding runs.
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

# Validate the committed clean PNG without decoding or re-encoding it.
$owlBytes = [System.IO.File]::ReadAllBytes($normalizedOwlPath)
$pngSignature = [byte[]](0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A)
if ($owlBytes.Length -lt $pngSignature.Length) {
    throw 'Normalized NeZnaika owl resource is too short to be a PNG.'
}
for ($i = 0; $i -lt $pngSignature.Length; $i++) {
    if ($owlBytes[$i] -ne $pngSignature[$i]) {
        throw 'Normalized NeZnaika owl resource does not have a valid PNG signature.'
    }
}

# IMPORTANT: remove the Image object from the WinForms resx completely. Keeping a
# System.Drawing.Bitmap in AboutBox.resx makes MSBuild GenerateResource serialize it
# through GDI+, which is exactly where the Windows runner was failing with MSB4018.
# The PNG itself is now a raw EmbeddedResource declared in Directory.Build.targets.
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

Write-Host ('NeZnaika branding applied; AboutBox Image removed from resx; raw PNG manifest resource validated (' + $owlBytes.Length + ' bytes).')
