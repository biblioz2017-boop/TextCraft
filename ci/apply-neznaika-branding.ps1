$ErrorActionPreference = 'Stop'

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$version = '1.0.14.0'
$displayVersion = '1.0.14'
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
    $assembly = [regex]::Replace($assembly, 'AssemblyVersion\("[^"]+"\)', 'AssemblyVersion("1.0.14.0")')
    $assembly = [regex]::Replace($assembly, 'AssemblyFileVersion\("[^"]+"\)', 'AssemblyFileVersion("1.0.14.0")')
    Write-Utf8Text $assemblyPath $assembly
}

$designerPath = 'Forge.Designer.cs'
if (Test-Path $designerPath) {
    $designer = Read-Utf8Text $designerPath
    $designer = $designer.Replace('this.ForgeTab.Label = "neZnaika";', 'this.ForgeTab.Label = "НеZнайка";')
    $designer = $designer.Replace('this.ForgeTab.Label = "TextCraft";', 'this.ForgeTab.Label = "НеZнайка";')
    $designer = $designer.Replace('Выбрать локальную языковую модель neZnaika.', 'Выбрать локальную языковую модель НеZнайка.')
    Write-Utf8Text $designerPath $designer
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
    $setup = [regex]::Replace($setup, '"ProductVersion" = "8:[^"]+"', '"ProductVersion" = "8:1.0.14"')
    $setup = $setup.Replace('"Name" = "8:TextCraft.WordAddIn"', '"Name" = "8:НеZнайка.WordAddIn"')
    $setup = $setup.Replace('"Name" = "8:neZnaika.WordAddIn"', '"Name" = "8:НеZнайка.WordAddIn"')
    Write-Utf8Text $setupPath $setup
}

New-Item -ItemType Directory -Force 'artifact' | Out-Null

$installer = @'
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host ''
Write-Host 'НеZнайка 1.0.14 — установка' -ForegroundColor Cyan
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
Write-Host 'НеZнайка 1.0.14 установлена.' -ForegroundColor Green
Write-Host 'Откройте Word и вкладку «НеZнайка».'
Read-Host 'Нажмите Enter для выхода'
'@
Write-Utf8Text 'artifact/Install-NeZnaika.ps1' $installer

$cmd = @'
@echo off
chcp 65001 >nul
cd /d "%~dp0"
title NeZnaika 1.0.14 Installer
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
НеZнайка 1.0.14 для Microsoft Word
=================================

ВАЖНО: для установки запускайте файл:
    00_INSTALL-NeZnaika.cmd

Не используйте Install-TextCraft.cmd из старой CI-обвязки — он оставлен только
для совместимости старого workflow.

Установка:
1. Закройте Microsoft Word.
2. Запустите 00_INSTALL-NeZnaika.cmd.
3. Дождитесь сообщения об успешной установке.
4. Запустите Word и откройте вкладку «НеZнайка».

Внутренние имена TextCraft.dll и TextCraft.vsto сохранены специально: их смена
может сломать VSTO-манифест. Пользовательское имя продукта — НеZнайка.
Версия — 1.0.14.
'@
Write-Utf8Text 'artifact/00_README-FIRST-NeZnaika.txt' $readme
Write-Utf8Text 'artifact/README-FIRST-NeZnaika.txt' $readme

Write-Host 'НеZнайка branding, version 1.0.14 and installer bundle prepared.'
