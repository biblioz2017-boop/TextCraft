$ErrorActionPreference = 'Stop'

# Keep this script strictly ASCII for Windows PowerShell 5.1.
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$path = 'GenerateUserControl.Science.cs'
$fullPath = (Resolve-Path $path).Path
$text = [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8)

if (-not $text.Contains('using System.Linq;')) {
    $anchor = 'using System.Collections.Generic;'
    if (-not $text.Contains($anchor)) {
        throw 'Could not locate using block in GenerateUserControl.Science.cs.'
    }

    $newline = if ($text.Contains("`r`n")) { "`r`n" } else { "`n" }
    $text = $text.Replace($anchor, $anchor + $newline + 'using System.Linq;')
    [System.IO.File]::WriteAllText($fullPath, $text, $utf8NoBom)
    Write-Host 'Added System.Linq for automatic RAG topic detection.'
} else {
    Write-Host 'System.Linq is already available for automatic RAG topic detection.'
}
