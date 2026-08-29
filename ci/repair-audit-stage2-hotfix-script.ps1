$ErrorActionPreference = 'Stop'

# Repair the idempotency guard before the stage-2 hotfix edits C# source.
# The old guard matched the method name anywhere in the file, including the call
# that the same script had just inserted. That prevented the helper method body
# from being added and caused CS0103 during compilation.
$path = 'ci\apply-audit-stage2-hotfix.ps1'
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$fullPath = (Resolve-Path $path).Path
$text = [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8)

$old = "if (-not `$panel.Contains('GenerateAuditReviewIssuesWithFallbackAsync')) {"
$new = "if (-not `$panel.Contains('private async Task<List<AuditReviewIssue>> GenerateAuditReviewIssuesWithFallbackAsync(')) {"

if ($text.Contains($old)) {
    $text = $text.Replace($old, $new)
    [System.IO.File]::WriteAllText($fullPath, $text, $utf8NoBom)
    Write-Host 'Repaired audit stage-2 hotfix method insertion guard.'
} elseif ($text.Contains($new)) {
    Write-Host 'Audit stage-2 hotfix method insertion guard is already repaired.'
} else {
    throw 'Could not locate audit stage-2 hotfix method insertion guard.'
}
