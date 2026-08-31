$ErrorActionPreference = 'Stop'

# Repair the stage-2 hotfix itself before it edits C# source.
# The build runs MSBuild and then devenv for the setup project, so the patch can be
# executed more than once in the same workspace. It must therefore be idempotent.
$path = 'ci\apply-audit-stage2-hotfix.ps1'
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$fullPath = (Resolve-Path $path).Path
$text = [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8)

# 1) The old guard matched the method name anywhere in the file, including a call.
$oldGuard = "if (-not `$panel.Contains('GenerateAuditReviewIssuesWithFallbackAsync')) {"
$newGuard = "if (-not `$panel.Contains('private async Task<List<AuditReviewIssue>> GenerateAuditReviewIssuesWithFallbackAsync(')) {"
if ($text.Contains($oldGuard)) {
    $text = $text.Replace($oldGuard, $newGuard)
}

# 2) The original replacement was too broad. On the second build pass it found the
# GenerateAuditReviewIssuesAsync call INSIDE GenerateAuditReviewIssuesWithFallbackAsync
# and rewrote it into a recursive self-call. Limit replacement to the one panel call
# that has currentText, _lastAuditReport and 20 as its arguments. If crash-safe mode
# has already replaced that call with ParseAuditReportFallback, leave it alone.
$oldBlock = @'
$oldCall = 'List<AuditReviewIssue> issues = await GenerateAuditReviewIssuesAsync('
$newCall = 'List<AuditReviewIssue> issues = await GenerateAuditReviewIssuesWithFallbackAsync('
if ($panel.Contains($oldCall)) {
    $panel = $panel.Replace($oldCall, $newCall)
} elseif (-not $panel.Contains($newCall)) {
    throw 'Audit review generation call not found.'
}
'@

$newBlock = @'
$oldCallPattern = 'List<AuditReviewIssue>\s+issues\s*=\s*await\s+GenerateAuditReviewIssuesAsync\(\s*currentText,\s*_lastAuditReport,\s*20\s*\);'
$newCallBlock = @"
List<AuditReviewIssue> issues = await GenerateAuditReviewIssuesWithFallbackAsync(
                    currentText,
                    _lastAuditReport,
                    20
                );
"@
$newCallPattern = 'List<AuditReviewIssue>\s+issues\s*=\s*await\s+GenerateAuditReviewIssuesWithFallbackAsync\(\s*currentText,\s*_lastAuditReport,\s*20\s*\);'

if ([regex]::IsMatch($panel, $oldCallPattern)) {
    $panel = [regex]::Replace($panel, $oldCallPattern, $newCallBlock.Trim(), 1)
} elseif ($panel.Contains('issues = ParseAuditReportFallback(')) {
    Write-Host 'Crash-safe local audit parser is already active; stage-2 call left unchanged.'
} elseif (-not [regex]::IsMatch($panel, $newCallPattern)) {
    throw 'Audit review generation call not found.'
}
'@

if ($text.Contains($oldBlock.Trim())) {
    $text = $text.Replace($oldBlock.Trim(), $newBlock.Trim())
} elseif (-not $text.Contains('$oldCallPattern = ')) {
    throw 'Could not locate audit stage-2 call replacement block.'
}

[System.IO.File]::WriteAllText($fullPath, $text, $utf8NoBom)
Write-Host 'Repaired audit stage-2 hotfix for idempotent multi-pass builds.'
