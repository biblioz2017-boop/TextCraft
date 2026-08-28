$ErrorActionPreference = 'Stop'

$ps1Path = [System.IO.Path]::GetFullPath('artifact/Install-NeZnaika.ps1')
if (-not [System.IO.File]::Exists($ps1Path)) {
    throw 'Install-NeZnaika.ps1 was not generated.'
}

# Windows PowerShell 5.1 treats UTF-8 without BOM as the active ANSI code page.
# Re-save the generated Russian installer as UTF-8 WITH BOM so -File parses it
# correctly on Windows 10/11 regardless of the system locale.
$text = [System.IO.File]::ReadAllText($ps1Path, [System.Text.Encoding]::UTF8)
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($ps1Path, $text, $utf8Bom)

# Keep the launcher strictly ASCII so cmd.exe never has to guess its encoding.
$cmd = @'
@echo off
cd /d "%~dp0"
title NeZnaika 1.0.15 Installer
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-NeZnaika.ps1"
if errorlevel 1 (
  echo.
  echo NeZnaika installation failed. See the error above.
  pause
)
'@

$ascii = [System.Text.Encoding]::ASCII
foreach ($name in @('00_INSTALL-NeZnaika.cmd', 'INSTALL-NeZnaika.cmd')) {
    $path = [System.IO.Path]::GetFullPath((Join-Path 'artifact' $name))
    [System.IO.File]::WriteAllText($path, $cmd, $ascii)
}

# Russian README files should also open correctly in legacy editors.
foreach ($name in @('00_README-FIRST-NeZnaika.txt', 'README-FIRST-NeZnaika.txt')) {
    $path = [System.IO.Path]::GetFullPath((Join-Path 'artifact' $name))
    if ([System.IO.File]::Exists($path)) {
        $readme = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
        [System.IO.File]::WriteAllText($path, $readme, $utf8Bom)
    }
}

Write-Host 'Verified installer encoding: PowerShell UTF-8 BOM, CMD ASCII.'
