$ErrorActionPreference = 'Stop'

# Keep this file strictly ASCII for Windows PowerShell 5.1.
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Read-Utf8Text([string]$Path) {
    return [System.IO.File]::ReadAllText((Resolve-Path $Path).Path, [System.Text.Encoding]::UTF8)
}

function Write-Utf8Text([string]$Path, [string]$Text) {
    [System.IO.File]::WriteAllText((Resolve-Path $Path).Path, $Text, $utf8NoBom)
}

function Get-NewLine([string]$Text) {
    if ($Text.Contains("`r`n")) { return "`r`n" }
    return "`n"
}

Write-Host 'Applying audit post-apply handle guard...'

# -----------------------------------------------------------------------------
# AuditFix: never create a new WinForms handle merely to restore button state or
# append a status message after Word edits. Also persist full exception details.
# -----------------------------------------------------------------------------
$fixPath = 'GenerateUserControl.AuditFix.cs'
$fix = Read-Utf8Text $fixPath
$nl = Get-NewLine $fix

$helperStart = $fix.IndexOf('        private static void TrySetAuditControlEnabled(Control control, bool enabled)')
$helperEnd = if ($helperStart -ge 0) { $fix.IndexOf('        private void SetAuditControlsBusy(bool busy)', $helperStart) } else { -1 }
if ($helperStart -lt 0 -or $helperEnd -le $helperStart) {
    throw 'TrySetAuditControlEnabled helper was not found.'
}

$helper = @'
        private static bool IsAuditControlHandleUsable(Control control)
        {
            return control != null &&
                !control.IsDisposed &&
                !control.Disposing &&
                control.IsHandleCreated;
        }

        private static void TrySetAuditControlEnabled(Control control, bool enabled)
        {
            if (!IsAuditControlHandleUsable(control))
                return;

            try
            {
                control.Enabled = enabled;
            }
            catch (System.ComponentModel.Win32Exception)
            {
            }
            catch (InvalidOperationException)
            {
            }
        }

'@
$fix = $fix.Substring(0, $helperStart) + $helper.Replace("`n", $nl) + $fix.Substring($helperEnd)

$noticeStart = $fix.IndexOf('        private void AppendAuditFixNotice(string text)')
$noticeEnd = if ($noticeStart -ge 0) { $fix.IndexOf('        private static string NormalizeEditText(string value)', $noticeStart) } else { -1 }
if ($noticeStart -lt 0 -or $noticeEnd -le $noticeStart) {
    throw 'AppendAuditFixNotice method was not found.'
}

$noticeAndLog = @'
        private static string GetAuditDiagnosticLogPath()
        {
            try
            {
                return System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(),
                    "NeZnaika-audit-errors.log"
                );
            }
            catch
            {
                return "NeZnaika-audit-errors.log";
            }
        }

        private static void WriteAuditDiagnostic(string stage, Exception ex)
        {
            try
            {
                int handles = -1;
                try
                {
                    using (System.Diagnostics.Process process = System.Diagnostics.Process.GetCurrentProcess())
                        handles = process.HandleCount;
                }
                catch
                {
                }

                string entry =
                    DateTime.Now.ToString("O") + Environment.NewLine +
                    "Stage: " + (stage ?? string.Empty) + Environment.NewLine +
                    "Process handles: " + handles + Environment.NewLine +
                    (ex == null ? "Exception: <null>" : ex.ToString()) +
                    Environment.NewLine +
                    new string('-', 72) + Environment.NewLine;

                System.IO.File.AppendAllText(
                    GetAuditDiagnosticLogPath(),
                    entry,
                    System.Text.Encoding.UTF8
                );
            }
            catch
            {
            }
        }

        private void AppendAuditFixNotice(string text)
        {
            if (!IsAuditControlHandleUsable(_responseTextBox))
                return;

            try
            {
                if (_responseTextBox.TextLength > 0)
                    _responseTextBox.AppendText("\r\n\r\n");
                _responseTextBox.AppendText(
                    "[\u041d\u0435Z\u043d\u0430\u0439\u043a\u0430 \u2014 \u0438\u0441\u043f\u0440\u0430\u0432\u043b\u0435\u043d\u0438\u0435 \u043f\u043e \u0430\u0443\u0434\u0438\u0442\u0443] " +
                    (text ?? string.Empty)
                );
                _responseTextBox.SelectionStart = _responseTextBox.TextLength;
                _responseTextBox.ScrollToCaret();
            }
            catch (System.ComponentModel.Win32Exception ex)
            {
                WriteAuditDiagnostic("AppendAuditFixNotice", ex);
            }
            catch (InvalidOperationException ex)
            {
                WriteAuditDiagnostic("AppendAuditFixNotice", ex);
            }
        }

'@
$fix = $fix.Substring(0, $noticeStart) + $noticeAndLog.Replace("`n", $nl) + $fix.Substring($noticeEnd)

# Add stack-trace logging to the main audit-fix catch without changing its behavior.
$runStart = $fix.IndexOf('        private async Task RunAuditFixAsync(bool singleEdit)')
$runEnd = if ($runStart -ge 0) { $fix.IndexOf('        private async Task<List<AuditEdit>> GenerateAuditEditsAsync(', $runStart) } else { -1 }
if ($runStart -lt 0 -or $runEnd -le $runStart) {
    throw 'RunAuditFixAsync method was not found.'
}
$run = $fix.Substring($runStart, $runEnd - $runStart)
$catchMarker = '                string context = "'
if (-not $run.Contains('WriteAuditDiagnostic(stage, ex);')) {
    $catchIndex = $run.IndexOf($catchMarker)
    if ($catchIndex -lt 0) {
        throw 'RunAuditFixAsync exception catch marker was not found.'
    }
    $lineStart = $run.LastIndexOf($nl, $catchIndex)
    if ($lineStart -lt 0) { $lineStart = 0 } else { $lineStart += $nl.Length }
    $run = $run.Insert($lineStart, '                WriteAuditDiagnostic(stage, ex);' + $nl)
}

# The response label is cleanup-only. Do not create its handle after Word edits.
$run = $run.Replace(
    'if (_responseLabel != null && !_responseLabel.IsDisposed && !_responseLabel.Disposing)',
    'if (IsAuditControlHandleUsable(_responseLabel))'
)
$fix = $fix.Substring(0, $runStart) + $run + $fix.Substring($runEnd)
Write-Utf8Text $fixPath $fix

# -----------------------------------------------------------------------------
# AuditPanel: after edits, update the data model first. Refresh controls only when
# their handles already exist. This prevents lazy CreateHandle calls after Word
# Track Changes has modified the document and focus/selection state.
# -----------------------------------------------------------------------------
$panelPath = 'GenerateUserControl.AuditPanel.cs'
$panel = Read-Utf8Text $panelPath
$pnl = Get-NewLine $panel

$buttonsStart = $panel.IndexOf('        private void UpdateAuditActionButtons()')
$buttonsEnd = if ($buttonsStart -ge 0) { $panel.IndexOf('        private void SetAuditReviewControlsEnabled(bool enabled)', $buttonsStart) } else { -1 }
if ($buttonsStart -lt 0 -or $buttonsEnd -le $buttonsStart) {
    throw 'UpdateAuditActionButtons method was not found.'
}

$buttonsMethod = @'
        private void UpdateAuditActionButtons()
        {
            bool hasSafeChecked = false;
            if (IsAuditControlHandleUsable(_auditIssueList))
            {
                try
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
                catch (System.ComponentModel.Win32Exception ex)
                {
                    WriteAuditDiagnostic("UpdateAuditActionButtons.CheckedItems", ex);
                }
                catch (InvalidOperationException ex)
                {
                    WriteAuditDiagnostic("UpdateAuditActionButtons.CheckedItems", ex);
                }
            }

            if (IsAuditControlHandleUsable(_auditApplySelectedButton))
            {
                try { _auditApplySelectedButton.Enabled = !_auditReviewBusy && hasSafeChecked; }
                catch (Exception ex) { WriteAuditDiagnostic("UpdateAuditActionButtons.Apply", ex); }
            }

            if (IsAuditControlHandleUsable(_auditSelectSafeButton))
            {
                bool hasPendingSafe = _auditReviewIssues.Any(
                    i => i != null && i.AutoApplicable && !i.Applied
                );
                try { _auditSelectSafeButton.Enabled = !_auditReviewBusy && hasPendingSafe; }
                catch (Exception ex) { WriteAuditDiagnostic("UpdateAuditActionButtons.Safe", ex); }
            }
        }

'@
$panel = $panel.Substring(0, $buttonsStart) + $buttonsMethod.Replace("`n", $pnl) + $panel.Substring($buttonsEnd)

$controlsStart = $panel.IndexOf('        private void SetAuditReviewControlsEnabled(bool enabled)')
$controlsEnd = if ($controlsStart -ge 0) { $panel.IndexOf('        private void ShowAuditReviewPanel()', $controlsStart) } else { -1 }
if ($controlsStart -lt 0 -or $controlsEnd -le $controlsStart) {
    throw 'SetAuditReviewControlsEnabled method was not found.'
}

$controlsMethod = @'
        private void SetAuditReviewControlsEnabled(bool enabled)
        {
            try
            {
                if (IsAuditControlHandleUsable(_auditIssueList))
                    _auditIssueList.Enabled = enabled;
                if (IsAuditControlHandleUsable(_auditGoToButton))
                    _auditGoToButton.Enabled = enabled &&
                        IsAuditControlHandleUsable(_auditIssueList) &&
                        _auditIssueList.SelectedItem != null;
                if (IsAuditControlHandleUsable(_auditClosePanelButton))
                    _auditClosePanelButton.Enabled = enabled;
            }
            catch (Exception ex)
            {
                WriteAuditDiagnostic("SetAuditReviewControlsEnabled", ex);
            }

            try { UpdateAuditActionButtons(); } catch (Exception ex) { WriteAuditDiagnostic("SetAuditReviewControlsEnabled.Actions", ex); }
        }

'@
$panel = $panel.Substring(0, $controlsStart) + $controlsMethod.Replace("`n", $pnl) + $panel.Substring($controlsEnd)

$markStart = $panel.IndexOf('        private void MarkAuditReviewEditsApplied(IEnumerable<AuditEdit> edits)')
$markEnd = if ($markStart -ge 0) { $panel.IndexOf('        private async Task BuildAuditReviewPanelAsync()', $markStart) } else { -1 }
if ($markStart -lt 0 -or $markEnd -le $markStart) {
    throw 'MarkAuditReviewEditsApplied method was not found.'
}

$markMethod = @'
        private void TryRefreshAuditReviewAfterApply(bool selectFirst)
        {
            if (!IsAuditControlHandleUsable(_auditReviewPanel))
                return;

            try
            {
                if (IsAuditControlHandleUsable(_auditIssueList))
                    RenderAuditReviewIssues(selectFirst);

                int safeCount = _auditReviewIssues.Count(
                    i => i != null && i.AutoApplicable && !i.Applied
                );

                if (IsAuditControlHandleUsable(_auditReviewHeader))
                {
                    _auditReviewHeader.Text =
                        "\u0417\u0430\u043c\u0435\u0447\u0430\u043d\u0438\u044f \u0430\u0443\u0434\u0438\u0442\u0430: " + _auditReviewIssues.Count +
                        " (\u0431\u0435\u0437\u043e\u043f\u0430\u0441\u043d\u044b\u0445 \u043e\u0441\u0442\u0430\u043b\u043e\u0441\u044c: " + safeCount + ")";
                }
            }
            catch (Exception ex)
            {
                WriteAuditDiagnostic("TryRefreshAuditReviewAfterApply", ex);
            }
        }

        private void MarkAuditReviewEditsApplied(IEnumerable<AuditEdit> edits)
        {
            if (edits == null)
                return;

            var appliedKeys = new HashSet<string>(StringComparer.Ordinal);
            foreach (AuditEdit edit in edits)
            {
                if (edit != null)
                    appliedKeys.Add((edit.FindText ?? string.Empty) + "\n" + (edit.Replacement ?? string.Empty));
            }

            foreach (AuditReviewIssue issue in _auditReviewIssues)
            {
                if (issue == null)
                    continue;

                string key = (issue.FindText ?? string.Empty) + "\n" + (issue.Replacement ?? string.Empty);
                if (appliedKeys.Contains(key))
                    issue.Applied = true;
            }

            TryRefreshAuditReviewAfterApply(true);
        }

'@
$panel = $panel.Substring(0, $markStart) + $markMethod.Replace("`n", $pnl) + $panel.Substring($markEnd)

# Add exact diagnostics to the panel-apply catch and replace direct post-apply UI
# refresh/header assignment with the handle-safe helper.
$applyStart = $panel.IndexOf('        private async void AuditApplySelectedButton_Click(object sender, EventArgs e)')
$applyEnd = if ($applyStart -ge 0) { $panel.IndexOf('        private void AuditSelectSafeButton_Click(', $applyStart) } else { -1 }
if ($applyStart -lt 0 -or $applyEnd -le $applyStart) {
    throw 'AuditApplySelectedButton_Click method was not found.'
}
$apply = $panel.Substring($applyStart, $applyEnd - $applyStart)

$refreshPattern = '(?s)\s*RenderAuditReviewIssues\(false\);\s*_auditReviewHeader\.Text\s*=\s*"[^"]*"\s*\+\s*applied\s*\+\s*\(skipped > 0 \? "[^"]*" \+ skipped : string\.Empty\);'
if ([regex]::IsMatch($apply, $refreshPattern)) {
    $apply = [regex]::Replace($apply, $refreshPattern, $pnl + '                TryRefreshAuditReviewAfterApply(false);', 1)
} elseif (-not $apply.Contains('TryRefreshAuditReviewAfterApply(false);')) {
    # Some build-time passes may already have changed formatting. Remove the two
    # direct operations independently if they are still present.
    $apply = $apply.Replace('                RenderAuditReviewIssues(false);' + $pnl, '                TryRefreshAuditReviewAfterApply(false);' + $pnl)
}

if (-not $apply.Contains('WriteAuditDiagnostic("AuditApplySelectedButton_Click", ex);')) {
    $catchPattern = '            catch \(Exception ex\)\s*\{\s*CommonUtils\.DisplayError\(ex\);\s*\}'
    $catchBlock = @'
            catch (Exception ex)
            {
                WriteAuditDiagnostic("AuditApplySelectedButton_Click", ex);
                CommonUtils.DisplayError(ex);
            }
'@
    if ([regex]::IsMatch($apply, $catchPattern)) {
        $apply = [regex]::Replace($apply, $catchPattern, $catchBlock.TrimEnd().Replace("`n", $pnl), 1)
    } else {
        throw 'Panel-apply exception catch was not found.'
    }
}
$panel = $panel.Substring(0, $applyStart) + $apply + $panel.Substring($applyEnd)

# Final invariant checks.
if (-not $fix.Contains('GetAuditDiagnosticLogPath')) {
    throw 'Audit diagnostic logger was not inserted.'
}
if (-not $fix.Contains('IsAuditControlHandleUsable')) {
    throw 'Audit handle usability guard was not inserted.'
}
if (-not $panel.Contains('TryRefreshAuditReviewAfterApply')) {
    throw 'Post-apply panel refresh guard was not inserted.'
}
if (-not $panel.Contains('WriteAuditDiagnostic("AuditApplySelectedButton_Click", ex);')) {
    throw 'Panel apply stack-trace logging was not inserted.'
}

Write-Utf8Text $panelPath $panel
Write-Host 'Audit post-apply handle guard and diagnostics applied successfully.'
