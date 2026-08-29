$ErrorActionPreference = 'Stop'

# Keep this script strictly ASCII for Windows PowerShell 5.1.
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
function Read-Utf8Text([string]$Path) {
    return [System.IO.File]::ReadAllText((Resolve-Path $Path).Path, [System.Text.Encoding]::UTF8)
}
function Write-Utf8Text([string]$Path, [string]$Text) {
    [System.IO.File]::WriteAllText((Resolve-Path $Path).Path, $Text, $utf8NoBom)
}

Write-Host 'Applying RAG cache v2 namespace...'

$path = 'RAGControl.Science.cs'
$text = Read-Utf8Text $path

if (-not $text.Contains('NeZnaika-RAG-cache-v2|')) {
    $pattern = '(?ms)(^\s*string\s+identity\s*=\s*\r?\n)(\s*)Path\.GetFullPath\(filePath\)\s*\+\s*"\|"\s*\+'
    $match = [regex]::Match($text, $pattern)
    if (-not $match.Success) {
        throw 'Could not locate persistent RAG cache identity with regex.'
    }

    $indent = $match.Groups[2].Value
    $replacement =
        $match.Groups[1].Value +
        $indent + '"NeZnaika-RAG-cache-v2|" +' + [Environment]::NewLine +
        $indent + 'Path.GetFullPath(filePath) + "|" +'

    $text = [regex]::Replace(
        $text,
        $pattern,
        [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $replacement },
        1
    )
}

Write-Utf8Text $path $text

$verify = Read-Utf8Text $path
if (-not $verify.Contains('NeZnaika-RAG-cache-v2|')) {
    throw 'RAG cache v2 namespace is missing.'
}

Write-Host 'RAG cache v2 namespace applied successfully.'
