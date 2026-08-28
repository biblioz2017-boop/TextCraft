$ErrorActionPreference = 'Stop'

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Read-Utf8Text([string]$Path) {
    $fullPath = (Resolve-Path $Path).Path
    return [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8)
}

function Write-Utf8Text([string]$Path, [string]$Text) {
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $dir = [System.IO.Path]::GetDirectoryName($fullPath)
    if ($dir -and -not (Test-Path $dir)) {
        New-Item -ItemType Directory -Force $dir | Out-Null
    }
    [System.IO.File]::WriteAllText($fullPath, $Text, $utf8NoBom)
}

Write-Host 'Applying neZnaika branding...'

# Keep internal VSTO filenames/assembly identity compatible with existing manifests,
# but change all user-facing product names in the MSI project.
$setupPath = 'OfficeAddInSetup/OfficeAddInSetup.vdproj'
if (Test-Path $setupPath) {
    $setup = Read-Utf8Text $setupPath
    $setup = $setup.Replace('"ProductName" = "8:TextCraft"', '"ProductName" = "8:neZnaika"')
    $setup = $setup.Replace('"Manufacturer" = "8:suncloudsmoon"', '"Manufacturer" = "8:neZnaika"')
    $setup = $setup.Replace('"Title" = "8:OfficeAddinSetup"', '"Title" = "8:neZnaika Setup"')
    $setup = $setup.Replace('"Subject" = "8:AI Tools"', '"Subject" = "8:neZnaika for Microsoft Word"')
    $setup = $setup.Replace('"Keywords" = "8:textcraft installer,textcraft addin,textcraft,craft"', '"Keywords" = "8:neznaika installer,word addin,local ai,rag"')
    $setup = $setup.Replace('"ARPCOMMENTS" = "8:Integrates AI tools into Microsoft® Word® (independently developed, not affiliated with Microsoft)"', '"ARPCOMMENTS" = "8:neZnaika local AI tools for Microsoft Word"')
    $setup = $setup.Replace('"Name" = "8:TextCraft.WordAddIn"', '"Name" = "8:neZnaika.WordAddIn"')
    Write-Utf8Text $setupPath $setup
}

# Create a one-click installer bundle directly in artifact. build-patched.ps1 later adds
# the compiled VSTO files and any MSI/EXE produced by Visual Studio to this same folder.
New-Item -ItemType Directory -Force 'artifact' | Out-Null

$installer = @'
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host ''
Write-Host 'neZnaika installer' -ForegroundColor Cyan
Write-Host '------------------'
Write-Host 'This installer will unblock files, trust the included CI certificate when present, and install the Word add-in.'
Write-Host ''

$word = Get-Process WINWORD -ErrorAction SilentlyContinue
if ($word) {
    Write-Host 'Microsoft Word is running. Please close Word before installation.' -ForegroundColor Yellow
    Read-Host 'Press Enter after Word is closed'
}

Get-ChildItem -Path $root -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
    try { Unblock-File -LiteralPath $_.FullName -ErrorAction Stop } catch { }
}

$cert = Get-ChildItem -Path $root -File -Filter '*.cer' -ErrorAction SilentlyContinue | Select-Object -First 1
if ($cert) {
    Write-Host 'Installing trust certificate for current user...'
    Import-Certificate -FilePath $cert.FullName -CertStoreLocation 'Cert:\CurrentUser\Root' | Out-Null
    Import-Certificate -FilePath $cert.FullName -CertStoreLocation 'Cert:\CurrentUser\TrustedPublisher' | Out-Null
}

$msi = Get-ChildItem -Path $root -File -Filter '*.msi' -ErrorAction SilentlyContinue | Select-Object -First 1
if ($msi) {
    Write-Host ('Installing MSI: ' + $msi.Name)
    $p = Start-Process msiexec.exe -ArgumentList @('/i', ('"' + $msi.FullName + '"'), '/passive') -Wait -PassThru
    if ($p.ExitCode -ne 0) { throw ('MSI installation failed with exit code ' + $p.ExitCode) }
} else {
    $vsto = Join-Path $root 'TextCraft.vsto'
    if (-not (Test-Path $vsto)) { throw 'TextCraft.vsto was not found in the package.' }
    Write-Host 'Launching VSTO installer...'
    Start-Process -FilePath $vsto -Wait
}

Write-Host ''
Write-Host 'neZnaika installation step completed.' -ForegroundColor Green
Write-Host 'Start Microsoft Word and open the neZnaika ribbon tab.'
Read-Host 'Press Enter to close'
'@
Write-Utf8Text 'artifact/Install-neZnaika.ps1' $installer

$cmd = @'
@echo off
cd /d "%~dp0"
title neZnaika Installer
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-neZnaika.ps1"
if errorlevel 1 (
  echo.
  echo Installation failed. See the message above.
  pause
)
'@
Write-Utf8Text 'artifact/INSTALL-neZnaika.cmd' $cmd

$readme = @'
neZnaika for Microsoft Word
===========================

Recommended installation:
1. Close Microsoft Word.
2. Double-click INSTALL-neZnaika.cmd.
3. Follow the installer prompts.
4. Start Word and open the neZnaika ribbon tab.

The package keeps the internal TextCraft.dll/TextCraft.vsto filenames for VSTO compatibility,
but the visible add-in/product name is neZnaika.

If an MSI is included, the one-click installer uses it automatically. Otherwise it installs
through the included VSTO manifest and trust certificate.
'@
Write-Utf8Text 'artifact/README-FIRST-neZnaika.txt' $readme

Write-Host 'neZnaika branding and installer bundle prepared.'
