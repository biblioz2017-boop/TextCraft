$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

$owlPath = Join-Path $root 'Assets\NeZnaikaOwl.jpg'
$hiddenOwlPath = Join-Path $root 'Assets\NeZnaikaOwl.jpg.ci-hidden'
$brandingScript = Join-Path $root 'ci\apply-neznaika-branding.ps1'
$resxPath = Join-Path $root 'AboutBox.resx'

if (-not (Test-Path $owlPath)) { throw 'NeZnaika owl image is missing.' }
if (-not (Test-Path $brandingScript)) { throw 'Branding script is missing.' }
if (-not (Test-Path $resxPath)) { throw 'AboutBox.resx is missing.' }

# The legacy branding script also edits the image resource. On Windows PowerShell
# 5.1 / x86 GDI+ Image.FromFile may report OutOfMemoryException even for a valid
# JPEG. Hide the image only while the legacy text/deployment branding runs, then
# encode the same JPEG with WPF's JPEG decoder and PNG encoder instead.
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

Add-Type -AssemblyName PresentationCore

$input = $null
$output = $null
try {
    $input = [System.IO.File]::OpenRead($owlPath)
    $decoder = New-Object System.Windows.Media.Imaging.JpegBitmapDecoder(
        $input,
        [System.Windows.Media.Imaging.BitmapCreateOptions]::PreservePixelFormat,
        [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    )
    if ($decoder.Frames.Count -lt 1) { throw 'The owl JPEG contains no image frame.' }

    $encoder = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
    $encoder.Frames.Add($decoder.Frames[0])
    $output = New-Object System.IO.MemoryStream
    $encoder.Save($output)
    $owlBase64 = [Convert]::ToBase64String($output.ToArray())
}
finally {
    if ($output) { $output.Dispose() }
    if ($input) { $input.Dispose() }
}

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

Write-Host 'NeZnaika branding applied; owl resource encoded with WPF as PNG.'
