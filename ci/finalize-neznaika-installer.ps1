$ErrorActionPreference = 'Stop'

$artifact = [System.IO.Path]::GetFullPath('artifact')
if (-not [System.IO.Directory]::Exists($artifact)) {
    throw 'Artifact directory does not exist.'
}

$vstoPath = Join-Path $artifact 'TextCraft.vsto'
if (-not (Test-Path -LiteralPath $vstoPath)) {
    throw 'TextCraft.vsto is missing from artifact.'
}

[xml]$deployment = [System.IO.File]::ReadAllText($vstoPath, [System.Text.Encoding]::UTF8)
$identity = $deployment.SelectSingleNode("//*[local-name()='assemblyIdentity' and @name='TextCraft.vsto']")
if ($null -eq $identity -or [string]::IsNullOrWhiteSpace($identity.version)) {
    throw 'Could not read deployment version from TextCraft.vsto.'
}
$version4 = [string]$identity.version
$version = $version4 -replace '\.0$', ''

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$ascii = [System.Text.Encoding]::ASCII

$installer = @'
param(
    [ValidateSet('Install', 'Repair', 'Diagnostics', 'Uninstall')]
    [string]$Action = 'Install'
)

$ErrorActionPreference = 'Stop'
$PackageRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogPath = Join-Path $env:TEMP ('NeZnaika-setup-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.log')
$ProductRoot = Join-Path $env:LOCALAPPDATA 'NeZnaika'
$CurrentFile = Join-Path $ProductRoot 'current.txt'

function Write-Log {
    param([string]$Text, [ConsoleColor]$Color = [ConsoleColor]::Gray)
    $line = ('[{0}] {1}' -f (Get-Date -Format 'HH:mm:ss'), $Text)
    Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
    Write-Host $Text -ForegroundColor $Color
}

function Get-DeploymentInfo {
    param([string]$Root)
    $vsto = Join-Path $Root 'TextCraft.vsto'
    if (-not (Test-Path -LiteralPath $vsto)) { throw 'Не найден TextCraft.vsto.' }
    [xml]$xml = Get-Content -LiteralPath $vsto -Raw
    $node = $xml.SelectSingleNode("//*[local-name()='assemblyIdentity' and @name='TextCraft.vsto']")
    if ($null -eq $node -or [string]::IsNullOrWhiteSpace($node.version)) {
        throw 'Не удалось определить версию VSTO-манифеста.'
    }
    $v4 = [string]$node.version
    $v = $v4 -replace '\.0$', ''
    return [PSCustomObject]@{ Version = $v; Version4 = $v4; Manifest = $vsto }
}

function Get-Sha256Base64 {
    param([string]$Path)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $stream = [System.IO.File]::OpenRead($Path)
    try { return [Convert]::ToBase64String($sha.ComputeHash($stream)) }
    finally { $stream.Dispose(); $sha.Dispose() }
}

function Test-ManifestIntegrity {
    param([string]$Root)
    $deploymentPath = Join-Path $Root 'TextCraft.vsto'
    $appManifestPath = Join-Path $Root 'TextCraft.dll.manifest'
    $dllPath = Join-Path $Root 'TextCraft.dll'
    foreach ($required in @($deploymentPath, $appManifestPath, $dllPath)) {
        if (-not (Test-Path -LiteralPath $required)) { throw ('Отсутствует файл: ' + $required) }
    }

    [xml]$deploymentXml = Get-Content -LiteralPath $deploymentPath -Raw
    $appDependency = $deploymentXml.SelectSingleNode("//*[local-name()='dependentAssembly' and @codebase='TextCraft.dll.manifest']")
    if ($null -eq $appDependency) { throw 'В TextCraft.vsto отсутствует ссылка на TextCraft.dll.manifest.' }
    if ([int64]$appDependency.size -ne (Get-Item -LiteralPath $appManifestPath).Length) {
        throw 'Размер TextCraft.dll.manifest не соответствует VSTO-манифесту.'
    }
    $appDigest = $appDependency.SelectSingleNode(".//*[local-name()='DigestValue']")
    if ($null -ne $appDigest -and (Get-Sha256Base64 $appManifestPath) -ne $appDigest.InnerText.Trim()) {
        throw 'SHA-256 TextCraft.dll.manifest не соответствует VSTO-манифесту.'
    }

    [xml]$appXml = Get-Content -LiteralPath $appManifestPath -Raw
    $nodes = @()
    $nodes += @($appXml.SelectNodes("//*[local-name()='dependentAssembly' and @dependencyType='install' and @codebase]"))
    $nodes += @($appXml.SelectNodes("//*[local-name()='file' and @name]"))
    $count = 0
    foreach ($node in $nodes) {
        $relative = if ($node.HasAttribute('codebase')) { $node.GetAttribute('codebase') } else { $node.GetAttribute('name') }
        if ([string]::IsNullOrWhiteSpace($relative)) { continue }
        $path = Join-Path $Root $relative
        if (-not (Test-Path -LiteralPath $path)) { throw ('Отсутствует runtime-файл: ' + $relative) }
        if ($node.HasAttribute('size') -and [int64]$node.GetAttribute('size') -ne (Get-Item -LiteralPath $path).Length) {
            throw ('Неверный размер runtime-файла: ' + $relative)
        }
        $digest = $node.SelectSingleNode(".//*[local-name()='DigestValue']")
        if ($null -ne $digest -and (Get-Sha256Base64 $path) -ne $digest.InnerText.Trim()) {
            throw ('Неверный SHA-256 runtime-файла: ' + $relative)
        }
        $count++
    }
    return $count
}

function Get-DotNetRelease {
    try {
        return [int](Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release -ErrorAction Stop).Release
    } catch { return 0 }
}

function Find-VstoInstaller {
    $candidates = New-Object System.Collections.Generic.List[string]
    if ($env:CommonProgramFiles) {
        $candidates.Add((Join-Path $env:CommonProgramFiles 'Microsoft Shared\VSTO\10.0\VSTOInstaller.exe'))
    }
    $cpf86 = ${env:CommonProgramFiles(x86)}
    if ($cpf86) {
        $candidates.Add((Join-Path $cpf86 'Microsoft Shared\VSTO\10.0\VSTOInstaller.exe'))
    }
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) { return $candidate }
    }
    try {
        $cmd = Get-Command VSTOInstaller.exe -ErrorAction Stop
        return $cmd.Source
    } catch { return $null }
}

function Get-WordPath {
    foreach ($key in @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE'
    )) {
        try {
            $value = (Get-ItemProperty -LiteralPath $key -ErrorAction Stop).'(default)'
            if ($value -and (Test-Path -LiteralPath $value)) { return $value }
        } catch { }
    }
    return $null
}

function Wait-ForWordToClose {
    $word = Get-Process WINWORD -ErrorAction SilentlyContinue
    if (-not $word) { return }
    Write-Log ''
    Write-Log 'Microsoft Word сейчас открыт.' Yellow
    Write-Log 'Сохраните документы и закройте все окна Word. Установщик не завершает Word принудительно.' Yellow
    [void](Read-Host 'После закрытия Word нажмите Enter')
    if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
        throw 'Word всё ещё запущен. Установка остановлена, чтобы не потерять документы.'
    }
}

function Import-PublisherCertificate {
    param([string]$Root)
    $certPath = Join-Path $Root 'NeZnaika-CI.cer'
    if (-not (Test-Path -LiteralPath $certPath)) { throw 'Не найден сертификат NeZnaika-CI.cer.' }
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certPath)
    foreach ($storeName in @('Root', 'TrustedPublisher')) {
        $storePath = 'Cert:\CurrentUser\' + $storeName
        $exists = Get-ChildItem $storePath -ErrorAction SilentlyContinue | Where-Object { $_.Thumbprint -eq $cert.Thumbprint } | Select-Object -First 1
        if (-not $exists) {
            Import-Certificate -FilePath $certPath -CertStoreLocation $storePath | Out-Null
        }
    }
    return $cert
}

function Copy-PayloadToStableFolder {
    param([string]$SourceRoot, [string]$Version)
    New-Item -ItemType Directory -Path $ProductRoot -Force | Out-Null
    $destination = Join-Path $ProductRoot $Version
    $sourceFull = [System.IO.Path]::GetFullPath($SourceRoot).TrimEnd('\')
    $destFull = [System.IO.Path]::GetFullPath($destination).TrimEnd('\')
    if ([string]::Equals($sourceFull, $destFull, [StringComparison]::OrdinalIgnoreCase)) { return $destination }

    $stage = $destination + '.__new'
    if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force }
    New-Item -ItemType Directory -Path $stage -Force | Out-Null
    Get-ChildItem -LiteralPath $SourceRoot -Force | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $stage -Recurse -Force
    }
    if (Test-Path -LiteralPath $destination) { Remove-Item -LiteralPath $destination -Recurse -Force }
    Move-Item -LiteralPath $stage -Destination $destination
    return $destination
}

function Invoke-Vsto {
    param([string]$Installer, [string]$Verb, [string]$ManifestPath, [switch]$IgnoreFailure)
    $uri = (New-Object System.Uri([System.IO.Path]::GetFullPath($ManifestPath))).AbsoluteUri
    $args = ('/{0} "{1}" /Silent' -f $Verb, $uri)
    $process = Start-Process -FilePath $Installer -ArgumentList $args -Wait -PassThru
    if ($process.ExitCode -eq 0) { return $true }
    if ($IgnoreFailure) {
        Write-Log ('VSTOInstaller /' + $Verb + ' вернул код ' + $process.ExitCode + '; продолжаю.') DarkYellow
        return $false
    }

    Write-Log ('Тихая установка вернула код ' + $process.ExitCode + '. Повторяю с окном VSTOInstaller.') Yellow
    $visibleArgs = ('/{0} "{1}"' -f $Verb, $uri)
    $visible = Start-Process -FilePath $Installer -ArgumentList $visibleArgs -Wait -PassThru
    if ($visible.ExitCode -ne 0) { throw ('VSTOInstaller завершился с кодом ' + $visible.ExitCode + '.') }
    return $true
}

function Show-Diagnostics {
    param([string]$Root)
    Write-Log '=== Диагностика НеZнайки ===' Cyan
    try {
        $info = Get-DeploymentInfo $Root
        Write-Log ('Версия пакета: ' + $info.Version) Green
        $count = Test-ManifestIntegrity $Root
        Write-Log ('Целостность пакета: OK, проверено runtime-файлов: ' + $count) Green
    } catch {
        Write-Log ('Целостность пакета: ОШИБКА — ' + $_.Exception.Message) Red
    }
    $release = Get-DotNetRelease
    if ($release -ge 533320) { Write-Log ('.NET Framework 4.8.1: OK (Release ' + $release + ')') Green }
    else { Write-Log ('.NET Framework 4.8.1: не обнаружен (Release ' + $release + ')') Red }

    $vstoInstaller = Find-VstoInstaller
    if ($vstoInstaller) { Write-Log ('VSTO Runtime: OK — ' + $vstoInstaller) Green }
    else { Write-Log 'VSTO Runtime: не обнаружен.' Red }

    $word = Get-WordPath
    if ($word) { Write-Log ('Microsoft Word: ' + $word) Green }
    else { Write-Log 'Microsoft Word: путь не найден в App Paths.' Yellow }

    if (Test-Path -LiteralPath $CurrentFile) {
        $installedRoot = (Get-Content -LiteralPath $CurrentFile -Raw).Trim()
        Write-Log ('Текущая папка установки: ' + $installedRoot) Cyan
    } else {
        Write-Log 'Установка через новый мастер ещё не зарегистрирована.' Yellow
    }
    Write-Log ('Журнал: ' + $LogPath) Cyan
}

try {
    New-Item -ItemType File -Path $LogPath -Force | Out-Null
    $package = Get-DeploymentInfo $PackageRoot
    Write-Host ''
    Write-Log ('НеZнайка ' + $package.Version + ' — установка для Microsoft Word') Cyan
    Write-Log '========================================================' DarkCyan

    if ($Action -eq 'Diagnostics') {
        Show-Diagnostics $PackageRoot
        [void](Read-Host 'Нажмите Enter для выхода')
        exit 0
    }

    if ($Action -eq 'Uninstall') {
        Wait-ForWordToClose
        $targetRoot = $null
        if (Test-Path -LiteralPath $CurrentFile) { $targetRoot = (Get-Content -LiteralPath $CurrentFile -Raw).Trim() }
        if (-not $targetRoot -or -not (Test-Path -LiteralPath (Join-Path $targetRoot 'TextCraft.vsto'))) { $targetRoot = $PackageRoot }
        $installer = Find-VstoInstaller
        if (-not $installer) { throw 'Не найден VSTO Runtime (VSTOInstaller.exe).' }
        Write-Log 'Удаляю надстройку из Word…' Cyan
        Invoke-Vsto -Installer $installer -Verb 'Uninstall' -ManifestPath (Join-Path $targetRoot 'TextCraft.vsto') | Out-Null
        if (Test-Path -LiteralPath $CurrentFile) { Remove-Item -LiteralPath $CurrentFile -Force -ErrorAction SilentlyContinue }
        if ($targetRoot -like ($ProductRoot + '*') -and (Test-Path -LiteralPath $targetRoot)) {
            Remove-Item -LiteralPath $targetRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
        Write-Log 'НеZнайка удалена. Сертификат пользователя оставлен, чтобы не нарушить доверие к другим установленным версиям.' Green
        [void](Read-Host 'Нажмите Enter для выхода')
        exit 0
    }

    Wait-ForWordToClose
    Write-Log '[1/5] Проверяю пакет…' Cyan
    $checked = Test-ManifestIntegrity $PackageRoot
    Write-Log ('Пакет цел: проверено runtime-файлов ' + $checked + '.') Green

    $dotNetRelease = Get-DotNetRelease
    if ($dotNetRelease -lt 533320) {
        throw 'Требуется .NET Framework 4.8.1. Установите его и повторите запуск.'
    }
    $vstoInstaller = Find-VstoInstaller
    if (-not $vstoInstaller) {
        throw 'Не найден Microsoft Visual Studio Tools for Office Runtime (VSTO Runtime). Установите VSTO Runtime и повторите запуск.'
    }

    Write-Log '[2/5] Подготавливаю постоянную папку установки…' Cyan
    Get-ChildItem -LiteralPath $PackageRoot -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
        try { Unblock-File -LiteralPath $_.FullName -ErrorAction Stop } catch { }
    }
    $installedRoot = Copy-PayloadToStableFolder -SourceRoot $PackageRoot -Version $package.Version
    Test-ManifestIntegrity $installedRoot | Out-Null
    Write-Log ('Файлы установлены в: ' + $installedRoot) Green

    Write-Log '[3/5] Устанавливаю сертификат издателя для текущего пользователя…' Cyan
    $cert = Import-PublisherCertificate $installedRoot
    Write-Log ('Сертификат доверен: ' + $cert.Thumbprint) Green

    Write-Log '[4/5] Регистрирую надстройку VSTO…' Cyan
    $installedManifest = Join-Path $installedRoot 'TextCraft.vsto'
    if ($Action -eq 'Repair') {
        Invoke-Vsto -Installer $vstoInstaller -Verb 'Uninstall' -ManifestPath $installedManifest -IgnoreFailure | Out-Null
    } elseif (Test-Path -LiteralPath $CurrentFile) {
        $oldRoot = (Get-Content -LiteralPath $CurrentFile -Raw).Trim()
        if ([string]::Equals($oldRoot, $installedRoot, [StringComparison]::OrdinalIgnoreCase)) {
            Invoke-Vsto -Installer $vstoInstaller -Verb 'Uninstall' -ManifestPath $installedManifest -IgnoreFailure | Out-Null
        }
    }
    Invoke-Vsto -Installer $vstoInstaller -Verb 'Install' -ManifestPath $installedManifest | Out-Null

    Write-Log '[5/5] Завершаю установку…' Cyan
    Set-Content -LiteralPath $CurrentFile -Value $installedRoot -Encoding UTF8
    Write-Log ('Готово. НеZнайка ' + $package.Version + ' установлена.') Green
    Write-Log 'Теперь откройте Word и перейдите на вкладку «НеZнайка».' Green
    Write-Log ('Журнал установки: ' + $LogPath) DarkGray
    [void](Read-Host 'Нажмите Enter, чтобы закрыть установщик')
    exit 0
}
catch {
    Write-Host ''
    Write-Log ('ОШИБКА: ' + $_.Exception.Message) Red
    Write-Log ('Журнал установки: ' + $LogPath) Yellow
    [void](Read-Host 'Нажмите Enter, чтобы закрыть установщик')
    exit 1
}
'@

$launchers = @{
    '00_INSTALL-NeZnaika.cmd' = 'Install'
    'INSTALL-NeZnaika.cmd' = 'Install'
    '01_REPAIR-NeZnaika.cmd' = 'Repair'
    '02_DIAGNOSTICS-NeZnaika.cmd' = 'Diagnostics'
    '03_UNINSTALL-NeZnaika.cmd' = 'Uninstall'
}

$installerPath = Join-Path $artifact 'Install-NeZnaika.ps1'
[System.IO.File]::WriteAllText($installerPath, $installer, $utf8Bom)

foreach ($entry in $launchers.GetEnumerator()) {
    $launcher = "@echo off`r`ncd /d `"%~dp0`"`r`ntitle NeZnaika Setup`r`npowershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File `"%~dp0Install-NeZnaika.ps1`" -Action $($entry.Value)`r`nif errorlevel 1 (`r`n  echo.`r`n  echo NeZnaika setup failed. See the message and log above.`r`n  pause`r`n)`r`n"
    [System.IO.File]::WriteAllText((Join-Path $artifact $entry.Key), $launcher, $ascii)
}

$readme = @"
НеZнайка $version для Microsoft Word
====================================

БЫСТРАЯ УСТАНОВКА
1. Полностью закройте Microsoft Word.
2. Запустите 00_INSTALL-NeZnaika.cmd двойным щелчком.
3. Установщик сам проверит целостность пакета, .NET Framework 4.8.1 и VSTO Runtime.
4. Файлы будут скопированы в постоянную папку %LOCALAPPDATA%\NeZnaika\$version.
5. Сертификат сборки будет добавлен только для текущего пользователя в Root и TrustedPublisher.
6. После сообщения «Готово» запустите Word и откройте вкладку «НеZнайка».

ДОПОЛНИТЕЛЬНО
01_REPAIR-NeZnaika.cmd      — повторная регистрация надстройки без изменения её функций.
02_DIAGNOSTICS-NeZnaika.cmd — проверка пакета, .NET, VSTO Runtime, Word и текущей установки.
03_UNINSTALL-NeZnaika.cmd   — удаление регистрации надстройки и её постоянной папки.

Установка выполняется для текущего пользователя и обычно не требует прав администратора.
Внутренние имена TextCraft.dll и TextCraft.vsto сохранены намеренно: они являются частью VSTO-идентичности.

Если установка завершится ошибкой, установщик покажет путь к журналу NeZnaika-setup-*.log в папке TEMP.
"@
foreach ($readmeName in @('00_README-FIRST-NeZnaika.txt', 'README-FIRST-NeZnaika.txt')) {
    [System.IO.File]::WriteAllText((Join-Path $artifact $readmeName), $readme, $utf8Bom)
}

$escapedPath = $installerPath.Replace("'", "''")
$parseCommand = '$tokens=$null;$errors=$null;[System.Management.Automation.Language.Parser]::ParseFile(''' + $escapedPath + ''',[ref]$tokens,[ref]$errors)|Out-Null;if($errors.Count -gt 0){$errors|ForEach-Object{Write-Host $_.Message};exit 1}else{exit 0}'
& powershell.exe -NoLogo -NoProfile -Command $parseCommand
if ($LASTEXITCODE -ne 0) {
    throw 'Final installer failed Windows PowerShell 5.1 parser validation.'
}

foreach ($name in $launchers.Keys) {
    $bytes = [System.IO.File]::ReadAllBytes((Join-Path $artifact $name))
    if (@($bytes | Where-Object { $_ -gt 127 }).Count -ne 0) {
        throw ($name + ' contains non-ASCII bytes.')
    }
}

Write-Host ('Convenient NeZnaika ' + $version + ' installer created and validated for Windows PowerShell 5.1.')
