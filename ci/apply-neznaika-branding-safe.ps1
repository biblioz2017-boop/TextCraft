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

# The clean owl is committed as an actual PNG file. CI never decodes or re-encodes
# an image: it only validates the PNG signature, Base64-encodes the exact bytes, and
# injects those bytes into the WinForms .resx resource.
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
$owlBase64 = [Convert]::ToBase64String($owlBytes)

$resx = [System.IO.File]::ReadAllText($resxPath, [System.Text.Encoding]::UTF8)
$logoPattern = '(?s)(<data name="logoPictureBox.Image"[^>]*>\s*<value>).*?(</value>\s*</data>)'
if (-not [regex]::IsMatch($resx, $logoPattern)) {
    throw 'Could not locate logoPictureBox.Image in AboutBox.resx.'
}
$resx = [regex]::Replace(
    $resx,
    $logoPattern,
    { param($m) $m.Groups[1].Value + "`r`n        " + $owlBase64 + "`r`n      " + $m.Groups[2].Value },
    1
)

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($resxPath, $resx, $utf8NoBom)

Write-Host ('NeZnaika branding applied; exact clean PNG owl injected (' + $owlBytes.Length + ' bytes).')
