$ErrorActionPreference = 'Stop'

# ASCII-only repair for Windows PowerShell 5.1.
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Repair-File([string]$Path) {
    $fullPath = (Resolve-Path $Path).Path
    $text = [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8)

    # ObjectDisposedException derives from InvalidOperationException. If both
    # catches are emitted in that order, C# reports CS0160. The broader
    # InvalidOperationException handler already covers the disposed-control case.
    $pattern = '(?ms)(\s*catch \(InvalidOperationException\)\s*\{\s*\})\s*catch \(ObjectDisposedException\)\s*\{\s*\}'
    $updated = [regex]::Replace($text, $pattern, '$1')

    if ($updated -ne $text) {
        [System.IO.File]::WriteAllText($fullPath, $updated, $utf8NoBom)
        Write-Host ('Repaired catch ordering in ' + $Path)
    } else {
        Write-Host ('No duplicate disposed catch remained in ' + $Path)
    }
}

Repair-File 'GenerateUserControl.AuditFix.cs'
Repair-File 'GenerateUserControl.AuditPanel.cs'

Write-Host 'Audit apply catch ordering repaired.'
