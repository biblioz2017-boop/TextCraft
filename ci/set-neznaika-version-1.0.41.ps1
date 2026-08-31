$ErrorActionPreference = 'Stop'
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Read-Utf8Text([string]$Path) {
    return [System.IO.File]::ReadAllText((Resolve-Path $Path).Path, [System.Text.Encoding]::UTF8)
}

function Write-Utf8Text([string]$Path, [string]$Text) {
    [System.IO.File]::WriteAllText((Resolve-Path $Path).Path, $Text, $utf8NoBom)
}

$assemblyPath = 'Properties/AssemblyInfo.cs'
$assembly = Read-Utf8Text $assemblyPath
$assembly = [regex]::Replace($assembly, 'AssemblyVersion\("[^"]+"\)', 'AssemblyVersion("1.0.41.0")')
$assembly = [regex]::Replace($assembly, 'AssemblyFileVersion\("[^"]+"\)', 'AssemblyFileVersion("1.0.41.0")')
Write-Utf8Text $assemblyPath $assembly

$projectPath = 'TextCraft.csproj'
$project = Read-Utf8Text $projectPath
$project = [regex]::Replace($project, '<ApplicationVersion>[^<]+</ApplicationVersion>', '<ApplicationVersion>1.0.41.0</ApplicationVersion>')
Write-Utf8Text $projectPath $project

$setupPath = 'OfficeAddInSetup\OfficeAddInSetup.vdproj'
$setup = Read-Utf8Text $setupPath
$setup = [regex]::Replace($setup, '"ProductCode" = "8:\{[^}]+\}"', '"ProductCode" = "8:{CC8766CF-9A38-4C66-ADD6-FF0D029918DB}"', 1)
$setup = [regex]::Replace($setup, '"PackageCode" = "8:\{[^}]+\}"', '"PackageCode" = "8:{5993ECBD-522E-4FEE-88F5-C014B0A31531}"', 1)
$setup = [regex]::Replace($setup, '"ProductVersion" = "8:[^"]+"', '"ProductVersion" = "8:1.0.41"', 1)
Write-Utf8Text $setupPath $setup

Write-Host 'NeZnaika package version set to 1.0.41.'
