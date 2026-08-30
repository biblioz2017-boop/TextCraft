$ErrorActionPreference = 'Stop'

$artifact = [System.IO.Path]::GetFullPath('artifact')
if (-not [System.IO.Directory]::Exists($artifact)) {
    throw 'Artifact directory does not exist.'
}

$ascii = [System.Text.Encoding]::ASCII

$installer = @'
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

function Wait-ForWordToClose {
    $word = Get-Process WINWORD -ErrorAction SilentlyContinue
    if (-not $word) { return }

    Write-Host ''
    Write-Host 'Microsoft Word is running.' -ForegroundColor Yellow
    Write-Host 'Please close all Word windows before continuing.'
    [void](Read-Host 'Press Enter after Word is closed')

    $word = Get-Process WINWORD -ErrorAction SilentlyContinue
    if ($word) {
        throw 'Microsoft Word is still running. Close Word and run the installer again.'
    }
}

try {
    Write-Host ''
    Write-Host 'NeZnaika 1.0.34 - Microsoft Word add-in setup' -ForegroundColor Cyan
    Write-Host '================================================'

    Wait-ForWordToClose

    Write-Host '[1/4] Unblocking package files...'
    Get-ChildItem -Path $root -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
        try { Unblock-File -LiteralPath $_.FullName -ErrorAction Stop } catch { }
    }

    Write-Host '[2/4] Installing publisher certificate...'
    $cert = Get-ChildItem -Path $root -File -Filter '*.cer' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($cert) {
        Import-Certificate -FilePath $cert.FullName -CertStoreLocation 'Cert:\CurrentUser\Root' | Out-Null
        Import-Certificate -FilePath $cert.FullName -CertStoreLocation 'Cert:\CurrentUser\TrustedPublisher' | Out-Null
    } else {
        Write-Host 'No .cer file found; continuing without certificate import.' -ForegroundColor Yellow
    }

    Write-Host '[3/4] Installing add-in...'
    $msi = Get-ChildItem -Path $root -File -Filter '*.msi' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($msi) {
        $arguments = @('/i', ('"' + $msi.FullName + '"'), '/passive', '/norestart')
        $process = Start-Process -FilePath 'msiexec.exe' -ArgumentList $arguments -Wait -PassThru
        if ($process.ExitCode -notin @(0, 3010)) {
            throw ('MSI installation failed with exit code ' + $process.ExitCode + '.')
        }
    } else {
        $vsto = Join-Path $root 'TextCraft.vsto'
        if (-not (Test-Path -LiteralPath $vsto)) {
            throw 'TextCraft.vsto was not found in the installation package.'
        }
        Start-Process -FilePath $vsto -Wait
    }

    Write-Host '[4/4] Done.'
    Write-Host ''
    Write-Host 'NeZnaika 1.0.34 was installed successfully.' -ForegroundColor Green
    Write-Host 'Open Microsoft Word and use the NeZnaika tab.'
    [void](Read-Host 'Press Enter to close this window')
    exit 0
}
catch {
    Write-Host ''
    Write-Host ('Installation failed: ' + $_.Exception.Message) -ForegroundColor Red
    [void](Read-Host 'Press Enter to close this window')
    exit 1
}
'@

$launcher = @'
@echo off
cd /d "%~dp0"
title NeZnaika 1.0.34 Installer
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-NeZnaika.ps1"
if errorlevel 1 (
  echo.
  echo NeZnaika installation failed. See the message above.
  pause
)
'@

$installerPath = Join-Path $artifact 'Install-NeZnaika.ps1'
[System.IO.File]::WriteAllText($installerPath, $installer, $ascii)

foreach ($name in @('00_INSTALL-NeZnaika.cmd', 'INSTALL-NeZnaika.cmd')) {
    [System.IO.File]::WriteAllText((Join-Path $artifact $name), $launcher, $ascii)
}

foreach ($name in @('Install-NeZnaika.ps1', '00_INSTALL-NeZnaika.cmd', 'INSTALL-NeZnaika.cmd')) {
    $path = Join-Path $artifact $name
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $nonAscii = @($bytes | Where-Object { $_ -gt 127 })
    if ($nonAscii.Count -ne 0) {
        throw ($name + ' contains non-ASCII bytes and is unsafe for Windows PowerShell 5.1.')
    }
}

$escapedPath = $installerPath.Replace("'", "''")
$parseCommand = '$tokens=$null;$errors=$null;[System.Management.Automation.Language.Parser]::ParseFile(''' + $escapedPath + ''',[ref]$tokens,[ref]$errors)|Out-Null;if($errors.Count -gt 0){$errors|ForEach-Object{Write-Host $_.Message};exit 1}else{exit 0}'
& powershell.exe -NoLogo -NoProfile -Command $parseCommand
if ($LASTEXITCODE -ne 0) {
    throw 'Final installer failed Windows PowerShell 5.1 parser validation.'
}

Write-Host 'Final installer is ASCII-only and parses successfully in Windows PowerShell 5.1.'
