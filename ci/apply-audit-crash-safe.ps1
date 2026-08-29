$ErrorActionPreference = 'Stop'

# Keep this script ASCII-only for Windows PowerShell 5.1.
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Read-Utf8Text([string]$Path) {
    return [System.IO.File]::ReadAllText((Resolve-Path $Path).Path, [System.Text.Encoding]::UTF8)
}

function Write-Utf8Text([string]$Path, [string]$Text) {
    [System.IO.File]::WriteAllText((Resolve-Path $Path).Path, $Text, $utf8NoBom)
}

Write-Host 'Applying crash-safe audit panel mode...'
$path = 'GenerateUserControl.AuditPanel.cs'
$panel = Read-Utf8Text $path

# The second streaming LLM request is temporarily removed from the automatic audit path.
# Stage 1 still produces the diagnostic report. The review panel is then populated with
# the bounded local parser inserted by apply-audit-stage2-hotfix.ps1. This isolates the
# Word process from the new second request while keeping the audit report and navigation.
$callPattern = 'List<AuditReviewIssue> issues = await GenerateAuditReviewIssuesWithFallbackAsync\(\s*currentText,\s*_lastAuditReport,\s*20\s*\);'
$localBlock = @'
List<AuditReviewIssue> issues;
                SetAuditReviewProgressPhase("\u043b\u043e\u043a\u0430\u043b\u044c\u043d\u044b\u0439 \u0431\u0435\u0437\u043e\u043f\u0430\u0441\u043d\u044b\u0439 \u0440\u0430\u0437\u0431\u043e\u0440 \u043e\u0442\u0447\u0435\u0442\u0430");
                await Task.Yield();
                issues = ParseAuditReportFallback(
                    _lastAuditReport,
                    currentText,
                    20
                );
                _auditReviewFallbackNotice =
                    "\u0411\u0435\u0437\u043e\u043f\u0430\u0441\u043d\u044b\u0439 \u0440\u0435\u0436\u0438\u043c: \u0432\u0442\u043e\u0440\u043e\u0439 LLM-\u0432\u044b\u0437\u043e\u0432 \u0432\u0440\u0435\u043c\u0435\u043d\u043d\u043e \u043e\u0442\u043a\u043b\u044e\u0447\u0435\u043d; \u043f\u0430\u043d\u0435\u043b\u044c \u043f\u043e\u0441\u0442\u0440\u043e\u0435\u043d\u0430 \u043b\u043e\u043a\u0430\u043b\u044c\u043d\u043e \u043f\u043e \u0433\u043e\u0442\u043e\u0432\u043e\u043c\u0443 \u043e\u0442\u0447\u0435\u0442\u0443.";
'@
if ([regex]::IsMatch($panel, $callPattern)) {
    $panel = [regex]::Replace($panel, $callPattern, $localBlock.Trim(), 1)
} elseif (-not $panel.Contains('issues = ParseAuditReportFallback(')) {
    throw 'Could not switch audit panel to local crash-safe parsing.'
}

# Do not present the ribbon Stop button as active during local stage 2.
$modelPattern = '(?s)\s*Forge\.SetModelActivity\(\s*true,\s*"[^"]*"\s*\);'
if ([regex]::IsMatch($panel, $modelPattern)) {
    $panel = [regex]::Replace($panel, $modelPattern, '', 1)
}

# Bound the deterministic fallback input so malformed/very large reports cannot make
# regex processing consume excessive memory inside WINWORD.EXE.
$guardAnchor = @'
            if (maxIssues <= 0 || string.IsNullOrWhiteSpace(auditReport) || string.IsNullOrWhiteSpace(currentText))
                return result;

            var seen = new HashSet<string>(StringComparer.Ordinal);
'@
$guardReplacement = @'
            if (maxIssues <= 0 || string.IsNullOrWhiteSpace(auditReport) || string.IsNullOrWhiteSpace(currentText))
                return result;

            const int maxFallbackCharacters = 24000;
            if (auditReport.Length > maxFallbackCharacters)
                auditReport = auditReport.Substring(0, maxFallbackCharacters);
            if (currentText.Length > maxFallbackCharacters)
                currentText = currentText.Substring(0, maxFallbackCharacters);

            var seen = new HashSet<string>(StringComparer.Ordinal);
'@
if ($panel.Contains($guardAnchor.Trim())) {
    $panel = $panel.Replace($guardAnchor.Trim(), $guardReplacement.Trim())
} elseif (-not $panel.Contains('const int maxFallbackCharacters = 24000;')) {
    throw 'Could not add bounds to local audit fallback parser.'
}

Write-Utf8Text $path $panel
Write-Host 'Crash-safe audit panel mode applied: stage 2 is local and bounded.'
