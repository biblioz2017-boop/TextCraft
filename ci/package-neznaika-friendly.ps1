$ErrorActionPreference = 'Stop'

$artifact = [System.IO.Path]::GetFullPath('artifact')
if (-not [System.IO.Directory]::Exists($artifact)) {
    throw 'Artifact directory does not exist.'
}

# The classic installer must behave like the pre-MSI package: do not expose an MSI
# in the user package. Install-NeZnaika.ps1 will therefore use TextCraft.vsto from
# this extracted package instead of installing a VSTO deployment under Program Files.
Get-ChildItem -LiteralPath $artifact -File -Filter '*.msi' -ErrorAction SilentlyContinue |
    Remove-Item -Force
Get-ChildItem -LiteralPath $artifact -File -Filter 'setup.exe' -ErrorAction SilentlyContinue |
    Remove-Item -Force

$resources = Join-Path $artifact 'resources'
if (Test-Path -LiteralPath $resources) {
    Remove-Item -LiteralPath $resources -Recurse -Force
}
New-Item -ItemType Directory -Path $resources | Out-Null

$rootLauncher = '00_INSTALL-NeZnaika.cmd'
Get-ChildItem -LiteralPath $artifact -Force | Where-Object {
    $_.Name -ne $rootLauncher -and $_.Name -ne 'resources'
} | ForEach-Object {
    Move-Item -LiteralPath $_.FullName -Destination $resources -Force
}

$launcher = @'
@echo off
cd /d "%~dp0"
title NeZnaika 1.0.42 Installer
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0resources\Install-NeZnaika.ps1"
if errorlevel 1 (
  echo.
  echo NeZnaika installation failed. See the message above.
  pause
)
'@
[System.IO.File]::WriteAllText((Join-Path $artifact $rootLauncher), $launcher, [System.Text.Encoding]::ASCII)

$readme = @'
NeZnaika 1.0.42
================

INSTALLATION
1. Close Microsoft Word.
2. Run ..\00_INSTALL-NeZnaika.cmd from the folder above.
3. Keep this resources folder next to the installer while installing.

This package intentionally uses the classic VSTO installation path.
The resources folder contains TextCraft.vsto, certificate, DLLs, language resources and runtime files.
Do not move or run individual files from this folder unless troubleshooting.
'@
[System.IO.File]::WriteAllText((Join-Path $resources 'README-FIRST-NeZnaika.txt'), $readme, [System.Text.Encoding]::ASCII)

$rootItems = @(Get-ChildItem -LiteralPath $artifact -Force)
if ($rootItems.Count -ne 2) {
    throw ('Friendly package root must contain exactly 2 items; found ' + $rootItems.Count + '.')
}
if (-not (Test-Path -LiteralPath (Join-Path $artifact $rootLauncher))) { throw 'Root installer launcher is missing.' }
if (-not (Test-Path -LiteralPath $resources -PathType Container)) { throw 'resources directory is missing.' }
foreach ($required in @('Install-NeZnaika.ps1','NeZnaika-CI.cer','TextCraft.vsto','TextCraft.dll')) {
    if (-not (Test-Path -LiteralPath (Join-Path $resources $required))) {
        throw ('Required packaged resource is missing: ' + $required)
    }
}
if (Get-ChildItem -LiteralPath $resources -File -Filter '*.msi' -ErrorAction SilentlyContinue) {
    throw 'MSI must not be present in the classic user package.'
}

Write-Host 'Friendly NeZnaika package created: classic root launcher + VSTO resources folder.'
