$ErrorActionPreference = 'Stop'

# Keep this script ASCII-only for Windows PowerShell 5.1.
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Read-Utf8Text([string]$Path) {
    return [System.IO.File]::ReadAllText((Resolve-Path $Path).Path, [System.Text.Encoding]::UTF8)
}

function Write-Utf8Text([string]$Path, [string]$Text) {
    [System.IO.File]::WriteAllText((Resolve-Path $Path).Path, $Text, $utf8NoBom)
}

Write-Host 'Applying crash-safe non-blocking audit panel mode...'
$path = 'GenerateUserControl.AuditPanel.cs'
$panel = Read-Utf8Text $path

# Stage 1 still produces the diagnostic report. Stage 2 must not make another
# streaming LLM request and must not run CPU/regex parsing on WINWORD's UI thread.
$callPattern = 'List<AuditReviewIssue> issues = await GenerateAuditReviewIssuesWithFallbackAsync\(\s*currentText,\s*_lastAuditReport,\s*20\s*\);'
$legacyLocalPattern = '(?s)List<AuditReviewIssue> issues;\s*SetAuditReviewProgressPhase\("[^\"]*"\);\s*await Task\.Yield\(\);\s*issues = ParseAuditReportFallback\(\s*_lastAuditReport,\s*currentText,\s*20\s*\);\s*_auditReviewFallbackNotice\s*=\s*"[^\"]*";'
$nonBlockingMarker = 'Task<List<AuditReviewIssue>> localParseTask = Task.Run('

$localBlock = @'
List<AuditReviewIssue> issues;
                SetAuditReviewProgressPhase("\u043b\u043e\u043a\u0430\u043b\u044c\u043d\u044b\u0439 \u0440\u0430\u0437\u0431\u043e\u0440: \u0444\u043e\u043d\u043e\u0432\u0430\u044f \u043e\u0431\u0440\u0430\u0431\u043e\u0442\u043a\u0430");
                string auditReportSnapshot = _lastAuditReport ?? string.Empty;
                string currentTextSnapshot = currentText ?? string.Empty;
                const int maxLocalParseCharacters = 12000;
                if (auditReportSnapshot.Length > maxLocalParseCharacters)
                    auditReportSnapshot = auditReportSnapshot.Substring(0, maxLocalParseCharacters);
                if (currentTextSnapshot.Length > maxLocalParseCharacters)
                    currentTextSnapshot = currentTextSnapshot.Substring(0, maxLocalParseCharacters);

                Task<List<AuditReviewIssue>> localParseTask = Task.Run(() =>
                {
                    List<AuditReviewIssue> issues = ParseAuditReportFallback(
                        auditReportSnapshot,
                        currentTextSnapshot,
                        20
                    );
                    return issues;
                });

                Task completedLocalParseTask = await Task.WhenAny(
                    localParseTask,
                    Task.Delay(5000)
                );

                if (completedLocalParseTask == localParseTask)
                {
                    issues = await localParseTask;
                    _auditReviewFallbackNotice =
                        "\u0411\u0435\u0437\u043e\u043f\u0430\u0441\u043d\u044b\u0439 \u0440\u0435\u0436\u0438\u043c: \u0432\u0442\u043e\u0440\u043e\u0439 LLM-\u0432\u044b\u0437\u043e\u0432 \u043e\u0442\u043a\u043b\u044e\u0447\u0435\u043d; \u043f\u0430\u043d\u0435\u043b\u044c \u043f\u043e\u0441\u0442\u0440\u043e\u0435\u043d\u0430 \u043b\u043e\u043a\u0430\u043b\u044c\u043d\u043e \u043f\u043e \u0433\u043e\u0442\u043e\u0432\u043e\u043c\u0443 \u043e\u0442\u0447\u0435\u0442\u0443.";
                }
                else
                {
                    issues = new List<AuditReviewIssue>();
                    _auditReviewFallbackNotice =
                        "\u041b\u043e\u043a\u0430\u043b\u044c\u043d\u044b\u0439 \u0440\u0430\u0437\u0431\u043e\u0440 \u043f\u0440\u0435\u0432\u044b\u0441\u0438\u043b 5 \u0441. \u041e\u0442\u0447\u0435\u0442 \u0430\u0443\u0434\u0438\u0442\u0430 \u0441\u043e\u0445\u0440\u0430\u043d\u0435\u043d; \u043f\u0430\u043d\u0435\u043b\u044c \u0437\u0430\u043c\u0435\u0447\u0430\u043d\u0438\u0439 \u043f\u0440\u043e\u043f\u0443\u0449\u0435\u043d\u0430, \u0447\u0442\u043e\u0431\u044b Word \u043e\u0441\u0442\u0430\u0432\u0430\u043b\u0441\u044f \u043e\u0442\u0437\u044b\u0432\u0447\u0438\u0432\u044b\u043c.";
                }
'@

if ([regex]::IsMatch($panel, $callPattern)) {
    $panel = [regex]::Replace($panel, $callPattern, $localBlock.Trim(), 1)
} elseif ([regex]::IsMatch($panel, $legacyLocalPattern)) {
    $panel = [regex]::Replace($panel, $legacyLocalPattern, $localBlock.Trim(), 1)
    Write-Host 'Upgraded legacy UI-thread local audit parsing.'
} elseif ($panel.Contains($nonBlockingMarker)) {
    Write-Host 'Non-blocking local audit parser is already active.'
} else {
    throw 'Could not switch audit panel to non-blocking local parsing.'
}

# Keep all second-stage UI work inside the BuildAuditReviewPanelAsync segment and
# remove only that segment's model-activity start marker. The second stage is local.
$buildStart = $panel.IndexOf('        private async Task BuildAuditReviewPanelAsync()')
if ($buildStart -lt 0) { throw 'BuildAuditReviewPanelAsync start not found.' }
$buildEnd = $panel.IndexOf('        private async Task<List<AuditReviewIssue>> GenerateAuditReviewIssuesWithFallbackAsync(', $buildStart)
if ($buildEnd -lt 0) {
    $buildEnd = $panel.IndexOf('        private async Task<List<AuditReviewIssue>> GenerateAuditReviewIssuesAsync(', $buildStart)
}
if ($buildEnd -lt 0) { throw 'BuildAuditReviewPanelAsync end not found.' }

$build = $panel.Substring($buildStart, $buildEnd - $buildStart)
$build = [regex]::Replace(
    $build,
    '\s*Forge\.SetModelActivity\(\s*true,\s*"[^\"]*"\s*\);',
    '',
    1
)
$build = [regex]::Replace(
    $build,
    '_auditReasonTextBox\.Text\s*=\s*"[^\"]*";',
    '_auditReasonTextBox.Text = "\u0420\u0430\u0437\u0431\u0438\u0440\u0430\u044e \u0433\u043e\u0442\u043e\u0432\u044b\u0439 \u043e\u0442\u0447\u0435\u0442 \u043b\u043e\u043a\u0430\u043b\u044c\u043d\u043e; \u0432\u0442\u043e\u0440\u043e\u0439 \u0432\u044b\u0437\u043e\u0432 LLM \u043d\u0435 \u0432\u044b\u043f\u043e\u043b\u043d\u044f\u0435\u0442\u0441\u044f.";',
    1
)
$build = [regex]::Replace(
    $build,
    '_responseLabel\.Text\s*=\s*"[^\"]*";',
    '_responseLabel.Text = "\u0410\u0443\u0434\u0438\u0442 \u2014 \u044d\u0442\u0430\u043f 2 \u0438\u0437 2: \u043b\u043e\u043a\u0430\u043b\u044c\u043d\u043e \u0441\u0442\u0440\u0443\u043a\u0442\u0443\u0440\u0438\u0440\u0443\u044e \u0437\u0430\u043c\u0435\u0447\u0430\u043d\u0438\u044f\u2026";',
    1
)
$panel = $panel.Substring(0, $buildStart) + $build + $panel.Substring($buildEnd)

# The progress timer must describe local work instead of claiming that an LLM stream
# is active. This also makes a future screenshot diagnostically useful.
$progressStart = $panel.IndexOf('        private void StartAuditReviewProgress()')
$progressEnd = $panel.IndexOf('        private void SetAuditReviewProgressPhase(', $progressStart)
if ($progressStart -lt 0 -or $progressEnd -lt 0) { throw 'Audit progress method not found.' }
$progress = $panel.Substring($progressStart, $progressEnd - $progressStart)
$progress = [regex]::Replace(
    $progress,
    '_auditReviewProgressPhase\s*=\s*"[^\"]*";',
    '_auditReviewProgressPhase = "\u043b\u043e\u043a\u0430\u043b\u044c\u043d\u044b\u0439 \u0440\u0430\u0437\u0431\u043e\u0440: \u043f\u043e\u0434\u0433\u043e\u0442\u043e\u0432\u043a\u0430";',
    1
)
$panel = $panel.Substring(0, $progressStart) + $progress + $panel.Substring($progressEnd)

# Bound parser input. Keep the existing 24k internal safety ceiling for compatibility,
# while the automatic UI path above passes at most 12k characters from each input.
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

# The only regex built from report text gets an explicit timeout. If it ever hits the
# timeout, the existing BuildAuditReviewPanelAsync exception handler reports the problem
# while Word remains responsive because parsing is already on a worker thread.
$unsafeDynamicMatch = '            Match match = Regex.Match(currentText, pattern, RegexOptions.Singleline);'
$safeDynamicMatch = '            Match match = Regex.Match(currentText, pattern, RegexOptions.Singleline, TimeSpan.FromMilliseconds(750));'
if ($panel.Contains($unsafeDynamicMatch)) {
    $panel = $panel.Replace($unsafeDynamicMatch, $safeDynamicMatch)
} elseif (-not $panel.Contains($safeDynamicMatch)) {
    throw 'Could not add timeout to dynamic audit fallback regex.'
}

Write-Utf8Text $path $panel
Write-Host 'Crash-safe audit mode applied: local stage 2 runs off the UI thread with a watchdog.'
