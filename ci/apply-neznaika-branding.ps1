$ErrorActionPreference = 'Stop'

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$version = '1.0.13.0'
$displayVersion = '1.0.13'
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

# Re-assert user-facing assembly branding after the PR workflow writes its legacy
# TextCraft 1.0.12 version immediately before compilation.
$assemblyPath = 'Properties/AssemblyInfo.cs'
if (Test-Path $assemblyPath) {
    $assembly = Read-Utf8Text $assemblyPath
    $assembly = [regex]::Replace($assembly, 'AssemblyTitle\("[^"]*"\)', 'AssemblyTitle("НеZнайка")')
    $assembly = [regex]::Replace($assembly, 'AssemblyProduct\("[^"]*"\)', 'AssemblyProduct("НеZнайка")')
    $assembly = [regex]::Replace($assembly, 'AssemblyTrademark\("[^"]*"\)', 'AssemblyTrademark("НеZнайка")')
    $assembly = [regex]::Replace($assembly, 'AssemblyVersion\("[^"]+"\)', 'AssemblyVersion("1.0.13.0")')
    $assembly = [regex]::Replace($assembly, 'AssemblyFileVersion\("[^"]+"\)', 'AssemblyFileVersion("1.0.13.0")')
    Write-Utf8Text $assemblyPath $assembly
}

# Keep the internal VSTO filenames/assembly identity compatible with existing manifests,
# but use the final product name everywhere the user sees it.
$designerPath = 'Forge.Designer.cs'
if (Test-Path $designerPath) {
    $designer = Read-Utf8Text $designerPath
    $designer = $designer.Replace('this.ForgeTab.Label = "neZnaika";', 'this.ForgeTab.Label = "НеZнайка";')
    $designer = $designer.Replace('this.ForgeTab.Label = "TextCraft";', 'this.ForgeTab.Label = "НеZнайка";')
    $designer = $designer.Replace('Выбрать локальную языковую модель neZnaika.', 'Выбрать локальную языковую модель НеZнайка.')
    Write-Utf8Text $designerPath $designer
}

$aboutPath = 'AboutBox.cs'
if (Test-Path $aboutPath) {
    $about = Read-Utf8Text $aboutPath
    $about = $about.Replace('О программе — neZnaika', 'О программе — НеZнайка')
    $about = $about.Replace('this.labelProductName.Text = "neZnaika";', 'this.labelProductName.Text = "НеZнайка";')
    $about = $about.Replace('this.labelVersion.Text = string.Format(_cultureHelper.GetLocalizedString("[AboutBox()] this.labelVersion.Text"), AssemblyVersion);', 'this.labelVersion.Text = "Версия " + AssemblyVersion;')
    Write-Utf8Text $aboutPath $about
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
    $setup = $setup.Replace('"Subject" = "8:neZnaika for Microsoft Word"', '"Subject" = "8:НеZнайка для Microsoft Word"')
    $setup = [regex]::Replace($setup, '"ProductVersion" = "8:[^"]+"', '"ProductVersion" = "8:1.0.13"')
    $setup = $setup.Replace('"Keywords" = "8:textcraft installer,textcraft addin,textcraft,craft"', '"Keywords" = "8:neznaika installer,word addin,local ai,rag"')
    $setup = $setup.Replace('"ARPCOMMENTS" = "8:Integrates AI tools into Microsoft® Word® (independently developed, not affiliated with Microsoft)"', '"ARPCOMMENTS" = "8:НеZнайка 1.0.13 — локальная AI-надстройка для Microsoft Word"')
    $setup = $setup.Replace('"ARPCOMMENTS" = "8:neZnaika local AI tools for Microsoft Word"', '"ARPCOMMENTS" = "8:НеZнайка 1.0.13 — локальная AI-надстройка для Microsoft Word"')
    $setup = $setup.Replace('"Name" = "8:TextCraft.WordAddIn"', '"Name" = "8:НеZнайка.WordAddIn"')
    $setup = $setup.Replace('"Name" = "8:neZnaika.WordAddIn"', '"Name" = "8:НеZнайка.WordAddIn"')
    Write-Utf8Text $setupPath $setup
}

New-Item -ItemType Directory -Force 'artifact' | Out-Null

$installer = @'
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host ''
Write-Host 'НеZнайка 1.0.13 — установка' -ForegroundColor Cyan
Write-Host '------------------------'
Write-Host 'Установщик разблокирует файлы, добавит сертификат доверия при наличии и установит надстройку Word.'
Write-Host ''

$word = Get-Process WINWORD -ErrorAction SilentlyContinue
if ($word) {
    Write-Host 'Microsoft Word запущен. Закройте Word перед установкой.' -ForegroundColor Yellow
    Read-Host 'После закрытия Word нажмите Enter'
}

Get-ChildItem -Path $root -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
    try { Unblock-File -LiteralPath $_.FullName -ErrorAction Stop } catch { }
}

$cert = Get-ChildItem -Path $root -File -Filter '*.cer' -ErrorAction SilentlyContinue | Select-Object -First 1
if ($cert) {
    Write-Host 'Устанавливаю сертификат доверия для текущего пользователя...'
    Import-Certificate -FilePath $cert.FullName -CertStoreLocation 'Cert:\CurrentUser\Root' | Out-Null
    Import-Certificate -FilePath $cert.FullName -CertStoreLocation 'Cert:\CurrentUser\TrustedPublisher' | Out-Null
}

$msi = Get-ChildItem -Path $root -File -Filter '*.msi' -ErrorAction SilentlyContinue | Select-Object -First 1
if ($msi) {
    Write-Host ('Устанавливаю MSI: ' + $msi.Name)
    $p = Start-Process msiexec.exe -ArgumentList @('/i', ('"' + $msi.FullName + '"'), '/passive') -Wait -PassThru
    if ($p.ExitCode -ne 0) { throw ('Ошибка MSI, код ' + $p.ExitCode) }
} else {
    $vsto = Join-Path $root 'TextCraft.vsto'
    if (-not (Test-Path $vsto)) { throw 'В пакете не найден TextCraft.vsto.' }
    Write-Host 'Запускаю VSTO-установку...'
    Start-Process -FilePath $vsto -Wait
}

Write-Host ''
Write-Host 'Установка НеZнайка 1.0.13 завершена.' -ForegroundColor Green
Write-Host 'Запустите Microsoft Word и откройте вкладку «НеZнайка».'
Read-Host 'Нажмите Enter для выхода'
'@
Write-Utf8Text 'artifact/Install-NeZnaika.ps1' $installer

$cmd = @'
@echo off
cd /d "%~dp0"
title NeZnaika 1.0.13 Installer
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-NeZnaika.ps1"
if errorlevel 1 (
  echo.
  echo Installation failed. See the message above.
  pause
)
'@
Write-Utf8Text 'artifact/INSTALL-NeZnaika.cmd' $cmd

$readme = @'
НеZнайка 1.0.13 для Microsoft Word
=================================

Рекомендуемая установка:
1. Закройте Microsoft Word.
2. Запустите INSTALL-NeZnaika.cmd.
3. Дождитесь завершения установки.
4. Запустите Word и откройте вкладку «НеZнайка».

Внутренние имена TextCraft.dll/TextCraft.vsto сохранены для совместимости VSTO.
Пользовательское имя продукта: НеZнайка.
Версия: 1.0.13.

Если в пакете есть MSI, установщик использует его автоматически. Если MSI отсутствует,
будет запущена установка через VSTO-манифест и сертификат доверия.
'@
Write-Utf8Text 'artifact/README-FIRST-NeZnaika.txt' $readme

Write-Host 'НеZнайка branding, version 1.0.13 and installer bundle prepared.'
