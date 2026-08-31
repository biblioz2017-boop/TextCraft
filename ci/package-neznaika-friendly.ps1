$ErrorActionPreference = 'Stop'

$artifact = [System.IO.Path]::GetFullPath('artifact')
if (-not [System.IO.Directory]::Exists($artifact)) {
    throw 'Artifact directory does not exist.'
}

$resources = Join-Path $artifact 'resources'
if (Test-Path -LiteralPath $resources) {
    Remove-Item -LiteralPath $resources -Recurse -Force
}
New-Item -ItemType Directory -Path $resources | Out-Null

# Keep only the obvious launcher in the archive root. Everything else is a resource.
$rootLauncher = '00_INSTALL-NeZnaika.cmd'
Get-ChildItem -LiteralPath $artifact -Force | Where-Object {
    $_.Name -ne $rootLauncher -and $_.Name -ne 'resources'
} | ForEach-Object {
    Move-Item -LiteralPath $_.FullName -Destination $resources -Force
}

$launcher = @'
@echo off
cd /d "%~dp0"
title NeZnaika 1.0.38 Installer
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0resources\Install-NeZnaika.ps1"
if errorlevel 1 (
  echo.
  echo NeZnaika installation failed. See the message above.
  pause
)
'@
[System.IO.File]::WriteAllText(
    (Join-Path $artifact $rootLauncher),
    $launcher,
    [System.Text.Encoding]::ASCII
)

$readme = @'
NeZnaika 1.0.38
================

INSTALLATION
1. Close Microsoft Word.
2. Run ..\00_INSTALL-NeZnaika.cmd from the folder above.
3. Keep this resources folder next to the installer while installing.

This folder contains the MSI, VSTO manifest, certificate, DLLs, language resources and other runtime files.
Do not run individual files from this folder unless troubleshooting.
'@
[System.IO.File]::WriteAllText(
    (Join-Path $resources 'README-FIRST-NeZnaika.txt'),
    $readme,
    [System.Text.Encoding]::ASCII
)

$rootItems = @(Get-ChildItem -LiteralPath $artifact -Force)
if ($rootItems.Count -ne 2) {
    throw ('Friendly package root must contain exactly 2 items; found ' + $rootItems.Count + '.')
}
if (-not (Test-Path -LiteralPath (Join-Path $artifact $rootLauncher))) {
    throw 'Root installer launcher is missing.'
}
if (-not (Test-Path -LiteralPath $resources -PathType Container)) {
    throw 'resources directory is missing.'
}
foreach ($required in @(
    'Install-NeZnaika.ps1',
    'NeZnaika-CI.cer',
    'TextCraft.vsto',
    'TextCraft.dll',
    'NeZnaika-1.0.38-Setup.msi'
)) {
    if (-not (Test-Path -LiteralPath (Join-Path $resources $required))) {
        throw ('Required packaged resource is missing: ' + $required)
    }
}

Write-Host 'Friendly NeZnaika package created: root launcher + resources folder.'
