$ErrorActionPreference = 'Stop'

$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
if (-not (Test-Path -LiteralPath $vswhere)) {
    throw 'vswhere.exe was not found.'
}

$vs = & $vswhere -latest -products * -property installationPath
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($vs)) {
    throw 'Visual Studio installation path could not be resolved.'
}

$expected = Join-Path $vs 'Common7\IDE\CommonExtensions\Microsoft\VSI\DisableOutOfProcBuild\DisableOutOfProcBuild.exe'
$tool = $null
if (Test-Path -LiteralPath $expected) {
    $tool = $expected
} else {
    $tool = Get-ChildItem -LiteralPath $vs -Recurse -File -Filter 'DisableOutOfProcBuild.exe' -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty FullName
}

if ([string]::IsNullOrWhiteSpace($tool) -or -not (Test-Path -LiteralPath $tool)) {
    throw 'DisableOutOfProcBuild.exe was not found. Visual Studio Installer Projects command-line builds require this registration step.'
}

$toolDirectory = Split-Path -Parent $tool
$toolName = Split-Path -Leaf $tool
Write-Host "Preparing Visual Studio Installer Projects command-line build with: $tool"

Push-Location $toolDirectory
try {
    & (Join-Path '.' $toolName)
    $exitCode = $LASTEXITCODE
}
finally {
    Pop-Location
}

if ($exitCode -ne 0) {
    throw "DisableOutOfProcBuild.exe failed with exit code $exitCode."
}

Write-Host 'Visual Studio Installer Projects out-of-process build has been disabled for the current CI user.'
