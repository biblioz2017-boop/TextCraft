$ErrorActionPreference = 'Stop'

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$version = '1.0.15.0'
$displayVersion = '1.0.15'
$productName = 'НеZнайка'

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

Write-Host "Applying $productName branding, version $displayVersion..."

$assemblyPath = 'Properties/AssemblyInfo.cs'
if (Test-Path $assemblyPath) {
    $assembly = Read-Utf8Text $assemblyPath
    $assembly = [regex]::Replace($assembly, 'AssemblyTitle\("[^"]*"\)', 'AssemblyTitle("НеZнайка")')
    $assembly = [regex]::Replace($assembly, 'AssemblyProduct\("[^"]*"\)', 'AssemblyProduct("НеZнайка")')
    $assembly = [regex]::Replace($assembly, 'AssemblyTrademark\("[^"]*"\)', 'AssemblyTrademark("НеZнайка")')
    $assembly = [regex]::Replace($assembly, 'AssemblyVersion\("[^"]+"\)', 'AssemblyVersion("1.0.15.0")')
    $assembly = [regex]::Replace($assembly, 'AssemblyFileVersion\("[^"]+"\)', 'AssemblyFileVersion("1.0.15.0")')
    Write-Utf8Text $assemblyPath $assembly
}

# Make the VSTO deployment metadata user-facing as НеZнайка while retaining the
# internal assembly/file names TextCraft.dll and TextCraft.vsto for compatibility.
$projectPath = 'TextCraft.csproj'
if (Test-Path $projectPath) {
    $project = Read-Utf8Text $projectPath
    $project = [regex]::Replace($project, '<ApplicationVersion>[^<]*</ApplicationVersion>', '<ApplicationVersion>1.0.15.0</ApplicationVersion>')
    $project = [regex]::Replace($project, '<ProductName>[^<]*</ProductName>', '<ProductName>НеZнайка</ProductName>')
    $project = [regex]::Replace($project, '<FriendlyName>[^<]*</FriendlyName>', '<FriendlyName>НеZнайка</FriendlyName>')
    $project = [regex]::Replace($project, '<OfficeApplicationDescription>[^<]*</OfficeApplicationDescription>', '<OfficeApplicationDescription>НеZнайка — локальная AI-надстройка для Microsoft Word</OfficeApplicationDescription>')
    Write-Utf8Text $projectPath $project
}

$designerPath = 'Forge.Designer.cs'
if (Test-Path $designerPath) {
    $designer = Read-Utf8Text $designerPath
    $designer = $designer.Replace('this.ForgeTab.Label = "neZnaika";', 'this.ForgeTab.Label = "НеZнайка";')
    $designer = $designer.Replace('this.ForgeTab.Label = "TextCraft";', 'this.ForgeTab.Label = "НеZнайка";')
    $designer = $designer.Replace('Выбрать локальную языковую модель neZnaika.', 'Выбрать локальную языковую модель НеZнайка.')
    Write-Utf8Text $designerPath $designer
}

# The About form used to be created on InitializeForge's worker STA thread and then
# shown on Word's UI thread. Windows Forms/GDI+ objects must not be moved between UI
# threads; this was a plausible source of ArgumentException("Недопустимый параметр").
# Create and dispose the dialog directly from the ribbon click handler instead.
$forgePath = 'Forge.cs'
if (Test-Path $forgePath) {
    $forge = Read-Utf8Text $forgePath
    $forge = [regex]::Replace(
        $forge,
        '(?m)^\s*_box = new AboutBox\(\);\s*\r?\n',
        ''
    )

    $aboutHandlerPattern = '(?s)        private void AboutButton_Click\(object sender, RibbonControlEventArgs e\)\s*\{.*?\n        \}\s*\n\s*        private void CancelButton_Click'
    $aboutHandlerReplacement = @'
        private void AboutButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                using (AboutBox box = new AboutBox())
                    box.ShowDialog();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void CancelButton_Click
'@
    if ([regex]::IsMatch($forge, $aboutHandlerPattern)) {
        $forge = [regex]::Replace($forge, $aboutHandlerPattern, $aboutHandlerReplacement, 1)
    } else {
        throw 'Could not locate AboutButton_Click for UI-thread-safe dialog creation.'
    }
    Write-Utf8Text $forgePath $forge
}

# Embed the exact user-provided owl-with-globe picture through AboutBox.resx. The
# existing WinForms Bitmap resource is PNG-encoded, so normalize the JPEG asset to
# PNG bytes before inserting it. Feeding raw JPEG bytes into this ResX Bitmap entry
# makes the Windows resource reader fail with MSB3103 / "Parameter is not valid".
$owlPath = 'Assets/NeZnaikaOwl.jpg'
$resxPath = 'AboutBox.resx'
if ((Test-Path $owlPath) -and (Test-Path $resxPath)) {
    Add-Type -AssemblyName System.Drawing
    $owlImage = $null
    $pngStream = $null
    try {
        $owlImage = [System.Drawing.Image]::FromFile((Resolve-Path $owlPath).Path)
        $pngStream = New-Object System.IO.MemoryStream
        $owlImage.Save($pngStream, [System.Drawing.Imaging.ImageFormat]::Png)
        $owlBase64 = [Convert]::ToBase64String($pngStream.ToArray())
    }
    finally {
        if ($pngStream) { $pngStream.Dispose() }
        if ($owlImage) { $owlImage.Dispose() }
    }

    $resx = Read-Utf8Text $resxPath
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
    Write-Utf8Text $resxPath $resx
    Write-Host 'Embedded verified owl-with-globe image into AboutBox.resx as PNG bitmap data.'
}

$setupPath = 'OfficeAddInSetup/OfficeAddInSetup.vdproj'
if (Test-Path $setupPath) {
    $setup = Read-Utf8Text $setupPath
    $setup = $setup.Replace('"ProductName" = "8:TextCraft"', '"ProductName" = "8:НеZнайка"')
    $setup = $setup.Replace('"ProductName" = "8:neZnaika"', '"ProductName" = "8:НеZнайка"')
    $setup = $setup.Replace('"Manufacturer" = "8:suncloudsmoon"', '"Manufacturer" = "8:НеZнайка"')
    $setup = $setup.Replace('"Manufacturer" = "8:neZnaika"', '"Manufacturer" = "8:НеZнайка"')
    $setup = $setup.Replace('"Title" = "8:OfficeAddinSetup"', '"Title" = "8:Установка НеZнайка"')
    $setup = $setup.Replace('"Title" = "8:neZnaika Setup"', '"Title" = "8:Установка НеZнайка"')
    $setup = $setup.Replace('"Subject" = "8:AI Tools"', '"Subject" = "8:НеZнайка для Microsoft Word"')
    $setup = [regex]::Replace($setup, '"ProductVersion" = "8:[^"]+"', '"ProductVersion" = "8:1.0.15"')
    $setup = $setup.Replace('"Name" = "8:TextCraft.WordAddIn"', '"Name" = "8:НеZнайка.WordAddIn"')
    $setup = $setup.Replace('"Name" = "8:neZnaika.WordAddIn"', '"Name" = "8:НеZнайка.WordAddIn"')
    Write-Utf8Text $setupPath $setup
}

New-Item -ItemType Directory -Force 'artifact' | Out-Null

$installer = @'
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host ''
Write-Host 'НеZнайка 1.0.15 — установка' -ForegroundColor Cyan
Write-Host '================================'
Write-Host 'Закройте Microsoft Word. Установщик сам разблокирует файлы, добавит сертификат и запустит установку.'
Write-Host ''

$word = Get-Process WINWORD -ErrorAction SilentlyContinue
if ($word) {
    Write-Host 'Microsoft Word сейчас запущен.' -ForegroundColor Yellow
    Read-Host 'Закройте Word и нажмите Enter'
}

Get-ChildItem -Path $root -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
    try { Unblock-File -LiteralPath $_.FullName -ErrorAction Stop } catch { }
}

$cert = Get-ChildItem -Path $root -File -Filter '*.cer' -ErrorAction SilentlyContinue | Select-Object -First 1
if ($cert) {
    Write-Host 'Добавляю сертификат доверия...'
    Import-Certificate -FilePath $cert.FullName -CertStoreLocation 'Cert:\CurrentUser\Root' | Out-Null
    Import-Certificate -FilePath $cert.FullName -CertStoreLocation 'Cert:\CurrentUser\TrustedPublisher' | Out-Null
}

$msi = Get-ChildItem -Path $root -File -Filter '*.msi' -ErrorAction SilentlyContinue | Select-Object -First 1
if ($msi) {
    Write-Host ('Запускаю НеZнайка MSI: ' + $msi.Name)
    $p = Start-Process msiexec.exe -ArgumentList @('/i', ('"' + $msi.FullName + '"'), '/passive', '/norestart') -Wait -PassThru
    if ($p.ExitCode -notin @(0, 3010)) { throw ('Ошибка MSI, код ' + $p.ExitCode) }
} else {
    $vsto = Join-Path $root 'TextCraft.vsto'
    if (-not (Test-Path $vsto)) { throw 'В пакете не найден VSTO-манифест.' }
    Write-Host 'MSI отсутствует, запускаю VSTO-установку...'
    Start-Process -FilePath $vsto -Wait
}

Write-Host ''
Write-Host 'НеZнайка 1.0.15 установлена.' -ForegroundColor Green
Write-Host 'Откройте Word и вкладку «НеZнайка».'
Read-Host 'Нажмите Enter для выхода'
'@
Write-Utf8Text 'artifact/Install-NeZnaika.ps1' $installer

$cmd = @'
@echo off
chcp 65001 >nul
cd /d "%~dp0"
title NeZnaika 1.0.15 Installer
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-NeZnaika.ps1"
if errorlevel 1 (
  echo.
  echo НеZнайка: установка завершилась с ошибкой.
  pause
)
'@
Write-Utf8Text 'artifact/00_INSTALL-NeZnaika.cmd' $cmd
Write-Utf8Text 'artifact/INSTALL-NeZnaika.cmd' $cmd

$readme = @'
НеZнайка 1.0.15 для Microsoft Word
=================================

ВАЖНО: для установки запускайте файл:
    00_INSTALL-NeZnaika.cmd

Установка:
1. Закройте Microsoft Word.
2. Запустите 00_INSTALL-NeZnaika.cmd.
3. Дождитесь сообщения об успешной установке.
4. Запустите Word и откройте вкладку «НеZнайка».

Внутренние имена TextCraft.dll и TextCraft.vsto сохранены специально: их смена
может сломать VSTO-манифест. Пользовательское имя продукта — НеZнайка.
Версия — 1.0.15.
'@
Write-Utf8Text 'artifact/00_README-FIRST-NeZnaika.txt' $readme
Write-Utf8Text 'artifact/README-FIRST-NeZnaika.txt' $readme

Write-Host 'НеZнайка branding, version 1.0.15, owl resource and installer bundle prepared.'
