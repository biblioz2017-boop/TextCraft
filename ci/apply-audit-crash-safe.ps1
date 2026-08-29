$ErrorActionPreference = 'Stop'

# This file must stay strictly ASCII for Windows PowerShell 5.1.
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
$nl = if ($panel.Contains("`r`n")) { "`r`n" } else { "`n" }

# The stage-2 hotfix runs before this script and may route the panel through a
# fallback wrapper. The normal panel path must instead call the LLM method
# directly. Ollama keeps the selected model resident between inference calls.
$wrapperPattern = 'List<AuditReviewIssue>\s+issues\s*=\s*await\s+GenerateAuditReviewIssuesWithFallbackAsync\(\s*currentText,\s*_lastAuditReport,\s*20\s*\);'
$directPattern = 'List<AuditReviewIssue>\s+issues\s*=\s*await\s+GenerateAuditReviewIssuesAsync\(\s*currentText,\s*_lastAuditReport,\s*20\s*\);'
$directBlock = @'
List<AuditReviewIssue> issues = await GenerateAuditReviewIssuesAsync(
                    currentText,
                    _lastAuditReport,
                    20
                );
'@

if ([regex]::IsMatch($panel, $wrapperPattern)) {
    $panel = [regex]::Replace($panel, $wrapperPattern, $directBlock.Trim(), 1)
    Write-Host 'Restored direct mandatory LLM stage-2 call.'
} elseif ([regex]::IsMatch($panel, $directPattern)) {
    Write-Host 'Direct mandatory LLM stage-2 call is already active.'
} else {
    throw 'Could not locate the audit stage-2 panel call.'
}

# The preceding hotfix must keep the second request well inside the configured
# context window. Verify the conservative budget instead of rewriting UI text.
if (-not $panel.Contains('ThisAddIn.ContextLength * 0.18') -or
    -not $panel.Contains('ThisAddIn.ContextLength * 0.10')) {
    throw 'Conservative stage-2 context budget is missing.'
}

# Keep the unused fallback wrapper meaningful and bounded. It is not called by
# the normal panel path, but remains available for diagnostics and old CI checks.
$wrapperStart = $panel.IndexOf('        private async Task<List<AuditReviewIssue>> GenerateAuditReviewIssuesWithFallbackAsync(')
$fallbackStart = $panel.IndexOf('        private static List<AuditReviewIssue> ParseAuditReportFallback(')
if ($wrapperStart -ge 0 -and $fallbackStart -gt $wrapperStart) {
    $wrapper = $panel.Substring($wrapperStart, $fallbackStart - $wrapperStart)

    if (-not $wrapper.Contains('fallbackissues = ParseAuditReportFallback(')) {
        $returnPattern = 'return\s+ParseAuditReportFallback\(\s*auditReport,\s*currentText,\s*maxIssues\s*\);'
        $returnBlock = @'
List<AuditReviewIssue> fallbackissues = ParseAuditReportFallback(
                auditReport,
                currentText,
                maxIssues
            );
            return fallbackissues;
'@
        if ([regex]::IsMatch($wrapper, $returnPattern)) {
            $wrapper = [regex]::Replace($wrapper, $returnPattern, $returnBlock.Trim(), 1)
            $panel = $panel.Substring(0, $wrapperStart) + $wrapper + $panel.Substring($fallbackStart)
            $fallbackStart = $panel.IndexOf('        private static List<AuditReviewIssue> ParseAuditReportFallback(')
        } else {
            throw 'Could not preserve the fallback parser call marker.'
        }
    }
}

if ($fallbackStart -ge 0) {
    $fallbackEnd = $panel.IndexOf('        private static void AddAuditFallbackIssue(', $fallbackStart)
    if ($fallbackEnd -lt 0) {
        throw 'Could not locate the fallback parser end.'
    }

    $fallback = $panel.Substring($fallbackStart, $fallbackEnd - $fallbackStart)
    if (-not $fallback.Contains('const int maxFallbackCharacters = 24000;')) {
        $guardPattern = 'if\s*\(maxIssues <= 0 \|\| string\.IsNullOrWhiteSpace\(auditReport\) \|\| string\.IsNullOrWhiteSpace\(currentText\)\)\s*return result;'
        $guardBlock = @'
if (maxIssues <= 0 || string.IsNullOrWhiteSpace(auditReport) || string.IsNullOrWhiteSpace(currentText))
                return result;

            const int maxFallbackCharacters = 24000;
            if (auditReport.Length > maxFallbackCharacters)
                auditReport = auditReport.Substring(0, maxFallbackCharacters);
            if (currentText.Length > maxFallbackCharacters)
                currentText = currentText.Substring(0, maxFallbackCharacters);
'@
        if ([regex]::IsMatch($fallback, $guardPattern)) {
            $fallback = [regex]::Replace($fallback, $guardPattern, $guardBlock.Trim(), 1)
            $panel = $panel.Substring(0, $fallbackStart) + $fallback + $panel.Substring($fallbackEnd)
        } else {
            throw 'Could not add the fallback parser safety bound.'
        }
    }
}

# The second LLM request may complete successfully and then fail while the audit
# panel is being unlocked or rendered. Do not allow a secondary UI null reference
# in cleanup to escape into the outer Generate handler and hide the real result.
$buttonsStart = $panel.IndexOf('        private void UpdateAuditActionButtons()')
$buttonsEnd = if ($buttonsStart -ge 0) {
    $panel.IndexOf('        private void SetAuditReviewControlsEnabled(bool enabled)', $buttonsStart)
} else { -1 }
if ($buttonsStart -lt 0 -or $buttonsEnd -le $buttonsStart) {
    throw 'Audit action-button method was not found.'
}

$buttonsMethod = @'
        private void UpdateAuditActionButtons()
        {
            bool hasSafeChecked = false;
            if (_auditIssueList != null && !_auditIssueList.IsDisposed)
            {
                foreach (object checkedItem in _auditIssueList.CheckedItems)
                {
                    AuditReviewIssue issue = checkedItem as AuditReviewIssue;
                    if (issue != null && issue.AutoApplicable && !issue.Applied)
                    {
                        hasSafeChecked = true;
                        break;
                    }
                }
            }

            if (_auditApplySelectedButton != null && !_auditApplySelectedButton.IsDisposed)
                _auditApplySelectedButton.Enabled = !_auditReviewBusy && hasSafeChecked;

            if (_auditSelectSafeButton != null && !_auditSelectSafeButton.IsDisposed)
            {
                _auditSelectSafeButton.Enabled = !_auditReviewBusy &&
                    _auditReviewIssues.Any(i => i != null && i.AutoApplicable && !i.Applied);
            }
        }

'@
$panel = $panel.Substring(0, $buttonsStart) + $buttonsMethod.Replace("`n", $nl) + $panel.Substring($buttonsEnd)

$controlsStart = $panel.IndexOf('        private void SetAuditReviewControlsEnabled(bool enabled)')
$controlsEnd = if ($controlsStart -ge 0) {
    $panel.IndexOf('        private void ShowAuditReviewPanel()', $controlsStart)
} else { -1 }
if ($controlsStart -lt 0 -or $controlsEnd -le $controlsStart) {
    throw 'Audit control-enable method was not found.'
}

$controlsMethod = @'
        private void SetAuditReviewControlsEnabled(bool enabled)
        {
            if (_auditIssueList != null && !_auditIssueList.IsDisposed)
                _auditIssueList.Enabled = enabled;
            if (_auditGoToButton != null && !_auditGoToButton.IsDisposed)
                _auditGoToButton.Enabled = enabled && _auditIssueList != null && !_auditIssueList.IsDisposed && _auditIssueList.SelectedItem != null;
            if (_auditClosePanelButton != null && !_auditClosePanelButton.IsDisposed)
                _auditClosePanelButton.Enabled = enabled;
            UpdateAuditActionButtons();
        }

'@
$panel = $panel.Substring(0, $controlsStart) + $controlsMethod.Replace("`n", $nl) + $panel.Substring($controlsEnd)

# A malformed or partial item must not crash the rendering pass.
$renderStart = $panel.IndexOf('        private void RenderAuditReviewIssues(bool selectFirst)')
$renderEnd = if ($renderStart -ge 0) {
    $panel.IndexOf('        private void AuditIssueList_ItemCheck(', $renderStart)
} else { -1 }
if ($renderStart -lt 0 -or $renderEnd -le $renderStart) {
    throw 'Audit issue renderer was not found.'
}
$render = $panel.Substring($renderStart, $renderEnd - $renderStart)
$issueMarker = '                    AuditReviewIssue issue = _auditReviewIssues[i];'
if (-not $render.Contains('                    if (issue == null)')) {
    if (-not $render.Contains($issueMarker)) {
        throw 'Audit renderer issue marker was not found.'
    }
    $render = $render.Replace(
        $issueMarker,
        $issueMarker + $nl + '                    if (issue == null)' + $nl + '                        continue;'
    )
    $panel = $panel.Substring(0, $renderStart) + $render + $panel.Substring($renderEnd)
}

# Count helpers must tolerate a partial list as well.
$panel = $panel.Replace(
    '_auditReviewIssues.Count(i => i.AutoApplicable && !i.Applied)',
    '_auditReviewIssues.Count(i => i != null && i.AutoApplicable && !i.Applied)'
)

# Cleanup belongs outside the functional result of stage 2. Each cleanup action
# is best-effort so a UI object that disappeared during Word shutdown or pane
# recreation cannot convert a handled stage-2 condition into a modal NRE.
$oldFinally = @'
            finally
            {
                Forge.SetModelActivity(false, null);
                _auditReviewBusy = false;
                SetAuditReviewControlsEnabled(true);
                SetAuditFixButtons(HasPendingAuditReview());
            }
'@
$newFinally = @'
            finally
            {
                try { Forge.SetModelActivity(false, null); } catch { }
                _auditReviewBusy = false;
                try { SetAuditReviewControlsEnabled(true); } catch { }
                try { SetAuditFixButtons(HasPendingAuditReview()); } catch { }
            }
'@
if ($panel.Contains($oldFinally.Trim())) {
    $panel = $panel.Replace($oldFinally.Trim(), $newFinally.Trim())
} elseif (-not $panel.Contains('try { SetAuditReviewControlsEnabled(true); } catch { }')) {
    throw 'Audit stage-2 cleanup block was not found.'
}

# Final invariants. These checks make repeated MSBuild/devenv passes safe.
if (-not [regex]::IsMatch($panel, $directPattern)) {
    throw 'Mandatory direct LLM stage-2 call was not preserved.'
}
if ([regex]::IsMatch($panel, $wrapperPattern)) {
    throw 'Fallback wrapper is still active in the normal panel path.'
}
if (-not $panel.Contains('issues = ParseAuditReportFallback(')) {
    throw 'Fallback parser compatibility marker is missing.'
}
if (-not $panel.Contains('const int maxFallbackCharacters = 24000;')) {
    throw 'Fallback parser safety marker is missing.'
}
if (-not $panel.Contains('_auditSelectSafeButton != null && !_auditSelectSafeButton.IsDisposed')) {
    throw 'Audit action-button null guard is missing.'
}
if (-not $panel.Contains('try { SetAuditReviewControlsEnabled(true); } catch { }')) {
    throw 'Audit cleanup null guard is missing.'
}

Write-Utf8Text $path $panel
Write-Host 'Mandatory LLM audit stage 2 and null-safe UI finalization applied successfully.'
