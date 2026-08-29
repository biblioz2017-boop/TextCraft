$ErrorActionPreference = 'Stop'

# Keep this script ASCII-only for Windows PowerShell 5.1.
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Read-Utf8Text([string]$Path) {
    return [System.IO.File]::ReadAllText((Resolve-Path $Path).Path, [System.Text.Encoding]::UTF8)
}

function Write-Utf8Text([string]$Path, [string]$Text) {
    [System.IO.File]::WriteAllText((Resolve-Path $Path).Path, $Text, $utf8NoBom)
}

Write-Host 'Applying mandatory LLM audit stage-2 mode...'
$path = 'GenerateUserControl.AuditPanel.cs'
$panel = Read-Utf8Text $path

# Stage 2 is mandatory LLM inference. Ollama keeps the already loaded model resident;
# this is a second inference request to the same ThisAddIn.Model and local endpoint,
# not a model unload/reload cycle.
$mandatoryCall = @'
List<AuditReviewIssue> issues = await GenerateAuditReviewIssuesAsync(
                    currentText,
                    _lastAuditReport,
                    20
                );
'@

$fallbackCallPattern = 'List<AuditReviewIssue>\s+issues\s*=\s*await\s+GenerateAuditReviewIssuesWithFallbackAsync\(\s*currentText,\s*_lastAuditReport,\s*20\s*\);'
$directCallPattern = 'List<AuditReviewIssue>\s+issues\s*=\s*await\s+GenerateAuditReviewIssuesAsync\(\s*currentText,\s*_lastAuditReport,\s*20\s*\);'
$localBlockPattern = '(?s)List<AuditReviewIssue> issues;\s*SetAuditReviewProgressPhase\([^;]+;.*?_auditReviewFallbackNotice\s*=.*?;\s*\}'

if ([regex]::IsMatch($panel, $fallbackCallPattern)) {
    $panel = [regex]::Replace($panel, $fallbackCallPattern, $mandatoryCall.Trim(), 1)
    Write-Host 'Replaced fallback stage-2 call with mandatory LLM call.'
} elseif ([regex]::IsMatch($panel, $localBlockPattern)) {
    $panel = [regex]::Replace($panel, $localBlockPattern, $mandatoryCall.Trim(), 1)
    Write-Host 'Removed local stage-2 parser and restored mandatory LLM call.'
} elseif ([regex]::IsMatch($panel, $directCallPattern)) {
    Write-Host 'Mandatory LLM stage-2 call is already active.'
} else {
    throw 'Could not locate audit stage-2 execution block.'
}

# Restore truthful LLM activity/status text in BuildAuditReviewPanelAsync.
$buildStart = $panel.IndexOf('        private async Task BuildAuditReviewPanelAsync()')
if ($buildStart -lt 0) { throw 'BuildAuditReviewPanelAsync start not found.' }
$buildEnd = $panel.IndexOf('        private async Task<List<AuditReviewIssue>> GenerateAuditReviewIssuesWithFallbackAsync(', $buildStart)
if ($buildEnd -lt 0) {
    $buildEnd = $panel.IndexOf('        private async Task<List<AuditReviewIssue>> GenerateAuditReviewIssuesAsync(', $buildStart)
}
if ($buildEnd -lt 0) { throw 'BuildAuditReviewPanelAsync end not found.' }

$build = $panel.Substring($buildStart, $buildEnd - $buildStart)
if (-not $build.Contains('Forge.SetModelActivity(true, "Структурирует аудит...");')) {
    $anchor = '                SetAuditFixButtons(false);'
    if (-not $build.Contains($anchor)) { throw 'Audit stage-2 activity anchor not found.' }
    $activityLine = '                Forge.SetModelActivity(true, "Структурирует аудит...");'
    $build = $build.Replace($anchor, $anchor + "`r`n" + $activityLine)
}
$build = [regex]::Replace(
    $build,
    '_auditReasonTextBox\.Text\s*=\s*"[^\"]*";',
    '_auditReasonTextBox.Text = "\u041f\u043e\u0434\u043e\u0436\u0434\u0438\u0442\u0435: LLM \u043f\u0440\u0435\u043e\u0431\u0440\u0430\u0437\u0443\u0435\u0442 \u043e\u0442\u0447\u0435\u0442 \u0432 \u043e\u0442\u0434\u0435\u043b\u044c\u043d\u044b\u0435 \u0437\u0430\u043c\u0435\u0447\u0430\u043d\u0438\u044f\u2026";',
    1
)
$build = [regex]::Replace(
    $build,
    '_responseLabel\.Text\s*=\s*"[^\"]*";',
    '_responseLabel.Text = "\u0410\u0443\u0434\u0438\u0442 \u2014 \u044d\u0442\u0430\u043f 2 \u0438\u0437 2: LLM \u0441\u0442\u0440\u0443\u043a\u0442\u0443\u0440\u0438\u0440\u0443\u0435\u0442 \u0437\u0430\u043c\u0435\u0447\u0430\u043d\u0438\u044f\u2026";',
    1
)
$panel = $panel.Substring(0, $buildStart) + $build + $panel.Substring($buildEnd)

# Progress starts by waiting for the first streamed token from the second inference.
$progressStart = $panel.IndexOf('        private void StartAuditReviewProgress()')
$progressEnd = $panel.IndexOf('        private void SetAuditReviewProgressPhase(', $progressStart)
if ($progressStart -lt 0 -or $progressEnd -lt 0) { throw 'Audit progress method not found.' }
$progress = $panel.Substring($progressStart, $progressEnd - $progressStart)
$progress = [regex]::Replace(
    $progress,
    '_auditReviewProgressPhase\s*=\s*"[^\"]*";',
    '_auditReviewProgressPhase = "LLM \u0430\u043a\u0442\u0438\u0432\u043d\u0430: \u043e\u0436\u0438\u0434\u0430\u044e \u043f\u0435\u0440\u0432\u044b\u0439 \u0444\u0440\u0430\u0433\u043c\u0435\u043d\u0442";',
    1
)
$panel = $panel.Substring(0, $progressStart) + $progress + $panel.Substring($progressEnd)

# Keep the conservative stage-2 context budget introduced by the hotfix.
$oldBudget = @'
            int textTokens = Math.Max(1200, (int)(ThisAddIn.ContextLength * 0.38));
            int auditTokens = Math.Max(800, (int)(ThisAddIn.ContextLength * 0.24));
'@
$newBudget = @'
            int textTokens = Math.Max(320, Math.Min(1600, (int)(ThisAddIn.ContextLength * 0.18)));
            int auditTokens = Math.Max(240, Math.Min(1000, (int)(ThisAddIn.ContextLength * 0.10)));
'@
if ($panel.Contains($oldBudget.Trim())) {
    $panel = $panel.Replace($oldBudget.Trim(), $newBudget.Trim())
} elseif (-not $panel.Contains('ThisAddIn.ContextLength * 0.18')) {
    throw 'Conservative audit stage-2 token budget is missing.'
}

# Guard against a runaway malformed model response while keeping streaming/progress.
if (-not $panel.Contains('const int maxStage2ResponseCharacters = 48000;')) {
    $appendPattern = 'response\.Append\(text\);\s*_auditReviewStreamedCharacters \+= text\.Length;'
    $appendReplacement = @'
const int maxStage2ResponseCharacters = 48000;
                    if (response.Length + text.Length > maxStage2ResponseCharacters)
                        throw new InvalidOperationException("LLM вернула слишком большой ответ на этапе 2 (> 48000 символов).");

                    response.Append(text);
                    _auditReviewStreamedCharacters += text.Length;
'@
    if ([regex]::IsMatch($panel, $appendPattern)) {
        $panel = [regex]::Replace($panel, $appendPattern, $appendReplacement.Trim(), 1)
    } else {
        throw 'Could not add stage-2 response size guard.'
    }
}

# Mandatory means mandatory: an empty/unparseable second-stage result is an error.
if (-not $panel.Contains('LLM не вернула пригодных структурированных замечаний')) {
    $returnPattern = 'SetAuditReviewProgressPhase\("проверяю привязку замечаний к тексту"\);\s*return ParseAuditReviewIssues\(response\.ToString\(\), currentText, maxIssues\);'
    $returnReplacement = @'
SetAuditReviewProgressPhase("проверяю привязку замечаний к тексту");
            List<AuditReviewIssue> parsedIssues = ParseAuditReviewIssues(response.ToString(), currentText, maxIssues);
            if (parsedIssues.Count == 0)
                throw new InvalidOperationException("LLM не вернула пригодных структурированных замечаний на этапе 2.");
            return parsedIssues;
'@
    if ([regex]::IsMatch($panel, $returnPattern)) {
        $panel = [regex]::Replace($panel, $returnPattern, $returnReplacement.Trim(), 1)
    } else {
        throw 'Could not enforce mandatory parsed stage-2 result.'
    }
}

Write-Utf8Text $path $panel
Write-Host 'Mandatory LLM audit stage 2 applied. Local parser is not used as the normal path.'
