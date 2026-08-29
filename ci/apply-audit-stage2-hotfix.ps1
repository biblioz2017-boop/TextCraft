$ErrorActionPreference = 'Stop'

# ASCII-only build-time hotfix for Windows PowerShell 5.1.
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Read-Utf8Text([string]$Path) {
    $fullPath = (Resolve-Path $Path).Path
    return [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8)
}

function Write-Utf8Text([string]$Path, [string]$Text) {
    $fullPath = (Resolve-Path $Path).Path
    [System.IO.File]::WriteAllText($fullPath, $Text, $utf8NoBom)
}

Write-Host 'Applying NeZnaika 1.0.19 audit stage-2 hotfix...'

$panelPath = 'GenerateUserControl.AuditPanel.cs'
$panel = Read-Utf8Text $panelPath

if (-not $panel.Contains('private string _auditReviewFallbackNotice = string.Empty;')) {
    $anchor = '        private string _auditReviewProgressPhase = string.Empty;'
    if (-not $panel.Contains($anchor)) { throw 'Audit progress field anchor not found.' }
    $panel = $panel.Replace(
        $anchor,
        $anchor + "`r`n        private string _auditReviewFallbackNotice = string.Empty;"
    )
}

$oldCall = 'List<AuditReviewIssue> issues = await GenerateAuditReviewIssuesAsync('
$newCall = 'List<AuditReviewIssue> issues = await GenerateAuditReviewIssuesWithFallbackAsync('
if ($panel.Contains($oldCall)) {
    $panel = $panel.Replace($oldCall, $newCall)
} elseif (-not $panel.Contains($newCall)) {
    throw 'Audit review generation call not found.'
}

$oldBudget = @'
            int textTokens = Math.Max(1200, (int)(ThisAddIn.ContextLength * 0.38));
            int auditTokens = Math.Max(800, (int)(ThisAddIn.ContextLength * 0.24));
'@
$newBudget = @'
            // Russian scientific text can consume substantially more tokens per character
            // than the old heuristic assumed. Keep stage 2 well below the advertised
            // context window because this request includes both text and the audit report.
            int textTokens = Math.Max(320, Math.Min(1600, (int)(ThisAddIn.ContextLength * 0.18)));
            int auditTokens = Math.Max(240, Math.Min(1000, (int)(ThisAddIn.ContextLength * 0.10)));
'@
if ($panel.Contains($oldBudget.Trim())) {
    $panel = $panel.Replace($oldBudget.Trim(), $newBudget.Trim())
} elseif (-not $panel.Contains('ThisAddIn.ContextLength * 0.18')) {
    throw 'Audit stage-2 token budget block not found.'
}

$methodAnchor = '        private async Task<List<AuditReviewIssue>> GenerateAuditReviewIssuesAsync('
if (-not $panel.Contains('GenerateAuditReviewIssuesWithFallbackAsync')) {
    if (-not $panel.Contains($methodAnchor)) { throw 'Audit stage-2 method anchor not found.' }

    $helper = @'
        private async Task<List<AuditReviewIssue>> GenerateAuditReviewIssuesWithFallbackAsync(
            string currentText,
            string auditReport,
            int maxIssues
        )
        {
            _auditReviewFallbackNotice = string.Empty;

            try
            {
                List<AuditReviewIssue> issues = await GenerateAuditReviewIssuesAsync(
                    currentText,
                    auditReport,
                    maxIssues
                );

                if (issues != null && issues.Count > 0)
                    return issues;

                _auditReviewFallbackNotice =
                    "\u0412\u0442\u043e\u0440\u043e\u0439 \u0432\u044b\u0437\u043e\u0432 LLM \u0437\u0430\u0432\u0435\u0440\u0448\u0438\u043b\u0441\u044f, \u043d\u043e \u043d\u0435 \u0432\u0435\u0440\u043d\u0443\u043b \u0440\u0430\u0441\u043f\u043e\u0437\u043d\u0430\u0432\u0430\u0435\u043c\u044b\u0445 \u0437\u0430\u043c\u0435\u0447\u0430\u043d\u0438\u0439. \u0418\u0441\u043f\u043e\u043b\u044c\u0437\u043e\u0432\u0430\u043d \u043b\u043e\u043a\u0430\u043b\u044c\u043d\u044b\u0439 \u0440\u0430\u0437\u0431\u043e\u0440 \u043e\u0442\u0447\u0435\u0442\u0430.";
                SetAuditReviewProgressPhase("\u043b\u043e\u043a\u0430\u043b\u044c\u043d\u044b\u0439 \u0440\u0435\u0437\u0435\u0440\u0432\u043d\u044b\u0439 \u0440\u0430\u0437\u0431\u043e\u0440");
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                string detail = ex.GetType().Name + ": " + (ex.Message ?? string.Empty);
                if (detail.Length > 260)
                    detail = detail.Substring(0, 260).TrimEnd() + "...";

                _auditReviewFallbackNotice =
                    "\u0412\u0442\u043e\u0440\u043e\u0439 \u0432\u044b\u0437\u043e\u0432 LLM \u043d\u0435 \u0443\u0434\u0430\u043b\u0441\u044f (" + detail +
                    "). \u041d\u0435Z\u043d\u0430\u0439\u043a\u0430 \u043f\u0435\u0440\u0435\u0448\u043b\u0430 \u043d\u0430 \u043b\u043e\u043a\u0430\u043b\u044c\u043d\u044b\u0439 \u0440\u0430\u0437\u0431\u043e\u0440 \u0433\u043e\u0442\u043e\u0432\u043e\u0433\u043e \u043e\u0442\u0447\u0435\u0442\u0430.";
                SetAuditReviewProgressPhase("\u043e\u0448\u0438\u0431\u043a\u0430 LLM; \u043b\u043e\u043a\u0430\u043b\u044c\u043d\u044b\u0439 \u0440\u0435\u0437\u0435\u0440\u0432\u043d\u044b\u0439 \u0440\u0430\u0437\u0431\u043e\u0440");
            }

            return ParseAuditReportFallback(auditReport, currentText, maxIssues);
        }

        private static List<AuditReviewIssue> ParseAuditReportFallback(
            string auditReport,
            string currentText,
            int maxIssues
        )
        {
            var result = new List<AuditReviewIssue>();
            if (maxIssues <= 0 || string.IsNullOrWhiteSpace(auditReport) || string.IsNullOrWhiteSpace(currentText))
                return result;

            var seen = new HashSet<string>(StringComparer.Ordinal);
            MatchCollection quoted = Regex.Matches(
                auditReport,
                "(?:\\u00AB(?<q1>.{12,360}?)\\u00BB|\\\"(?<q2>.{12,360}?)\\\"|\\u201C(?<q3>.{12,360}?)\\u201D)",
                RegexOptions.Singleline
            );

            foreach (Match match in quoted)
            {
                if (result.Count >= maxIssues)
                    break;

                string candidate = match.Groups["q1"].Success
                    ? match.Groups["q1"].Value
                    : (match.Groups["q2"].Success ? match.Groups["q2"].Value : match.Groups["q3"].Value);

                AddAuditFallbackIssue(result, seen, currentText, candidate, ExtractAuditFallbackReason(auditReport, match.Index), maxIssues);
            }

            if (result.Count >= maxIssues)
                return result;

            string normalizedAudit = Regex.Replace(auditReport, @"\s+", " ");
            string[] sentences = Regex.Split(currentText, @"(?<=[\.!?])\s+|[\r\n]+")
                .Select(value => NormalizeEditText(value))
                .Where(value => value.Length >= 24 && value.Length <= 360)
                .ToArray();

            foreach (string sentence in sentences)
            {
                if (result.Count >= maxIssues)
                    break;
                if (seen.Contains(sentence))
                    continue;
                if (normalizedAudit.IndexOf(sentence, StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                int reportIndex = normalizedAudit.IndexOf(sentence, StringComparison.OrdinalIgnoreCase);
                AddAuditFallbackIssue(
                    result,
                    seen,
                    currentText,
                    sentence,
                    ExtractAuditFallbackReason(normalizedAudit, reportIndex),
                    maxIssues
                );
            }

            return result;
        }

        private static void AddAuditFallbackIssue(
            List<AuditReviewIssue> result,
            HashSet<string> seen,
            string currentText,
            string candidate,
            string reason,
            int maxIssues
        )
        {
            if (result == null || result.Count >= maxIssues)
                return;

            string find = FindAuditFallbackSnippet(currentText, candidate);
            if (string.IsNullOrWhiteSpace(find) || !seen.Add(find))
                return;

            result.Add(new AuditReviewIssue
            {
                Category = "\u0420\u0443\u0447\u043d\u0430\u044f \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0430",
                FindText = find,
                Replacement = string.Empty,
                Reason = string.IsNullOrWhiteSpace(reason)
                    ? "\u0424\u0440\u0430\u0433\u043c\u0435\u043d\u0442 \u0443\u043f\u043e\u043c\u044f\u043d\u0443\u0442 \u0432 \u0434\u0438\u0430\u0433\u043d\u043e\u0441\u0442\u0438\u0447\u0435\u0441\u043a\u043e\u043c \u043e\u0442\u0447\u0435\u0442\u0435; \u0442\u0440\u0435\u0431\u0443\u0435\u0442 \u0440\u0443\u0447\u043d\u043e\u0439 \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0438."
                    : reason,
                ModelMarkedSafe = false,
                AutoApplicable = false,
                Applied = false
            });
        }

        private static string FindAuditFallbackSnippet(string currentText, string candidate)
        {
            string normalized = NormalizeEditText(candidate);
            if (normalized.Length < 12 || normalized.Length > 360)
                return null;

            int direct = currentText.IndexOf(normalized, StringComparison.Ordinal);
            if (direct >= 0)
                return currentText.Substring(direct, normalized.Length);

            string[] parts = Regex.Split(normalized, @"\s+")
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .ToArray();
            if (parts.Length < 2)
                return null;

            string pattern = string.Join(@"\s+", parts.Select(Regex.Escape));
            Match match = Regex.Match(currentText, pattern, RegexOptions.Singleline);
            return match.Success ? match.Value : null;
        }

        private static string ExtractAuditFallbackReason(string report, int index)
        {
            if (string.IsNullOrWhiteSpace(report))
                return string.Empty;

            int safeIndex = Math.Max(0, Math.Min(index, report.Length - 1));
            int start = report.LastIndexOf('\n', safeIndex);
            start = start < 0 ? 0 : start + 1;
            int end = report.IndexOf('\n', safeIndex);
            end = end < 0 ? report.Length : end;

            string reason = NormalizeEditText(report.Substring(start, Math.Max(0, end - start)));
            if (reason.Length > 320)
                reason = reason.Substring(0, 320).TrimEnd() + "...";
            return reason;
        }

'@

    $panel = $panel.Replace($methodAnchor, $helper + $methodAnchor)
}

$completeAnchor = '                CompleteAuditReviewProgress(_auditReviewIssues.Count);'
$fallbackProgress = @'
                CompleteAuditReviewProgress(_auditReviewIssues.Count);
                if (!string.IsNullOrWhiteSpace(_auditReviewFallbackNotice) && _auditReviewProgressLabel != null)
                    _auditReviewProgressLabel.Text += Environment.NewLine + _auditReviewFallbackNotice;
'@
if ($panel.Contains($completeAnchor) -and -not $panel.Contains('_auditReviewProgressLabel.Text += Environment.NewLine + _auditReviewFallbackNotice;')) {
    $panel = $panel.Replace($completeAnchor, $fallbackProgress.TrimEnd())
}

Write-Utf8Text $panelPath $panel

# Build 1.0.19 as a distinct installable package while preserving UpgradeCode.
$assemblyPath = 'Properties/AssemblyInfo.cs'
$assembly = Read-Utf8Text $assemblyPath
$assembly = [regex]::Replace($assembly, 'AssemblyVersion\("[^"]+"\)', 'AssemblyVersion("1.0.19.0")')
$assembly = [regex]::Replace($assembly, 'AssemblyFileVersion\("[^"]+"\)', 'AssemblyFileVersion("1.0.19.0")')
Write-Utf8Text $assemblyPath $assembly

$projectPath = 'TextCraft.csproj'
$project = Read-Utf8Text $projectPath
$project = [regex]::Replace($project, '<ApplicationVersion>[^<]+</ApplicationVersion>', '<ApplicationVersion>1.0.19.0</ApplicationVersion>')
Write-Utf8Text $projectPath $project

$setupPath = 'OfficeAddInSetup\OfficeAddInSetup.vdproj'
$setup = Read-Utf8Text $setupPath
$setup = [regex]::Replace($setup, '"ProductCode" = "8:\{[^}]+\}"', '"ProductCode" = "8:{24B9AA2D-313F-4E9F-B1C5-85042D7F8D33}"', 1)
$setup = [regex]::Replace($setup, '"PackageCode" = "8:\{[^}]+\}"', '"PackageCode" = "8:{9D1C3B4E-A85F-45C9-8B11-6B58A234C7EF}"', 1)
$setup = [regex]::Replace($setup, '"ProductVersion" = "8:[^"]+"', '"ProductVersion" = "8:1.0.19"', 1)
Write-Utf8Text $setupPath $setup

$finalizerPath = 'ci\finalize-neznaika-installer.ps1'
$finalizer = Read-Utf8Text $finalizerPath
$finalizer = $finalizer.Replace('1.0.18', '1.0.19')
Write-Utf8Text $finalizerPath $finalizer

Write-Host 'NeZnaika 1.0.19 audit stage-2 hotfix applied.'
