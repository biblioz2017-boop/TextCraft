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

Write-Host 'Applying audit apply UI-handle stability patch...'

# -----------------------------------------------------------------------------
# AuditFix: applying edits is the functional operation. UI cleanup and scrolling
# must never turn a successful Word edit into a Win32 window-handle exception.
# -----------------------------------------------------------------------------
$fixPath = 'GenerateUserControl.AuditFix.cs'
$fix = Read-Utf8Text $fixPath
$fixNl = Get-NewLine $fix

$busyStart = $fix.IndexOf('        private void SetAuditControlsBusy(bool busy)')
$busyEnd = if ($busyStart -ge 0) { $fix.IndexOf('        private void SetAuditFixButtons(bool enabled)', $busyStart) } else { -1 }
if ($busyStart -lt 0 -or $busyEnd -le $busyStart) {
    throw 'SetAuditControlsBusy method was not found.'
}

$busyMethod = @'
        private static void TrySetAuditControlEnabled(Control control, bool enabled)
        {
            if (control == null || control.IsDisposed || control.Disposing)
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
            catch (ObjectDisposedException)
            {
            }
        }

        private void SetAuditControlsBusy(bool busy)
        {
            TrySetAuditControlEnabled(_auditChapterButton, !busy);
            TrySetAuditControlEnabled(
                _fixAuditButton,
                !busy && !string.IsNullOrWhiteSpace(_lastAuditReport)
            );
            TrySetAuditControlEnabled(
                _nextAuditFixButton,
                !busy && !string.IsNullOrWhiteSpace(_lastAuditReport)
            );
            TrySetAuditControlEnabled(
                _resetAuditButton,
                !busy && _auditTargetRange != null
            );
            TrySetAuditControlEnabled(GenerateButton, !busy);
        }

'@
$fix = $fix.Substring(0, $busyStart) + $busyMethod.Replace("`n", $fixNl) + $fix.Substring($busyEnd)

$noticeStart = $fix.IndexOf('        private void AppendAuditFixNotice(string text)')
$noticeEnd = if ($noticeStart -ge 0) { $fix.IndexOf('        private static string NormalizeEditText(string value)', $noticeStart) } else { -1 }
if ($noticeStart -lt 0 -or $noticeEnd -le $noticeStart) {
    throw 'AppendAuditFixNotice method was not found.'
}

$noticeMethod = @'
        private void AppendAuditFixNotice(string text)
        {
            if (_responseTextBox == null || _responseTextBox.IsDisposed || _responseTextBox.Disposing)
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

                if (_responseTextBox.IsHandleCreated)
                {
                    try
                    {
                        _responseTextBox.ScrollToCaret();
                    }
                    catch (System.ComponentModel.Win32Exception)
                    {
                    }
                    catch (InvalidOperationException)
                    {
                    }
                }
            }
            catch (System.ComponentModel.Win32Exception)
            {
            }
            catch (InvalidOperationException)
            {
            }
            catch (ObjectDisposedException)
            {
            }
        }

'@
$fix = $fix.Substring(0, $noticeStart) + $noticeMethod.Replace("`n", $fixNl) + $fix.Substring($noticeEnd)

$runStart = $fix.IndexOf('        private async Task RunAuditFixAsync(bool singleEdit)')
$runEnd = if ($runStart -ge 0) { $fix.IndexOf('        private async Task<List<AuditEdit>> GenerateAuditEditsAsync(', $runStart) } else { -1 }
if ($runStart -lt 0 -or $runEnd -le $runStart) {
    throw 'RunAuditFixAsync method was not found.'
}
$run = $fix.Substring($runStart, $runEnd - $runStart)

$oldCleanupPattern = '(?s)            finally\s*\{\s*if \(modelActivityStarted\)\s*Forge\.SetModelActivity\(false, null\);\s*_auditFixBusy = false;\s*SetAuditControlsBusy\(false\);\s*if \(_responseLabel != null\)\s*_responseLabel\.Text = "[^"]*";\s*\}'
$newCleanup = @'
            finally
            {
                if (modelActivityStarted)
                {
                    try { Forge.SetModelActivity(false, null); } catch { }
                }

                _auditFixBusy = false;
                try { SetAuditControlsBusy(false); } catch { }

                try
                {
                    if (_responseLabel != null && !_responseLabel.IsDisposed && !_responseLabel.Disposing)
                        _responseLabel.Text = "\u0414\u0438\u0430\u043b\u043e\u0433 / \u043e\u0442\u0432\u0435\u0442:";
                }
                catch (System.ComponentModel.Win32Exception)
                {
                }
                catch (InvalidOperationException)
                {
                }
                catch (ObjectDisposedException)
                {
                }
            }
'@
if ([regex]::IsMatch($run, $oldCleanupPattern)) {
    $run = [regex]::Replace($run, $oldCleanupPattern, $newCleanup.TrimEnd().Replace("`n", $fixNl), 1)
} elseif (-not $run.Contains('try { SetAuditControlsBusy(false); } catch { }')) {
    throw 'RunAuditFixAsync cleanup block was not found.'
}
$fix = $fix.Substring(0, $runStart) + $run + $fix.Substring($runEnd)

Write-Utf8Text $fixPath $fix

# -----------------------------------------------------------------------------
# AuditPanel: programmatic CheckedListBox refresh used to fire ItemCheck for every
# issue and queue BeginInvoke calls. Suppress those callbacks during rendering and
# never force creation of a new Win32 handle merely to refresh button state.
# -----------------------------------------------------------------------------
$panelPath = 'GenerateUserControl.AuditPanel.cs'
$panel = Read-Utf8Text $panelPath
$panelNl = Get-NewLine $panel

if (-not $panel.Contains('private bool _auditRenderingIssues;')) {
    $fieldAnchor = '        private bool _auditReviewBusy;'
    if (-not $panel.Contains($fieldAnchor)) {
        throw 'Audit review busy field was not found.'
    }
    $panel = $panel.Replace(
        $fieldAnchor,
        $fieldAnchor + $panelNl + '        private bool _auditRenderingIssues;'
    )
}

$renderStart = $panel.IndexOf('        private void RenderAuditReviewIssues(bool selectFirst)')
$renderEnd = if ($renderStart -ge 0) { $panel.IndexOf('        private void AuditIssueList_ItemCheck(', $renderStart) } else { -1 }
if ($renderStart -lt 0 -or $renderEnd -le $renderStart) {
    throw 'RenderAuditReviewIssues method was not found.'
}

$renderMethod = @'
        private void RenderAuditReviewIssues(bool selectFirst)
        {
            if (_auditIssueList == null || _auditIssueList.IsDisposed || _auditIssueList.Disposing)
                return;

            bool beganUpdate = false;
            _auditRenderingIssues = true;
            try
            {
                try
                {
                    _auditIssueList.BeginUpdate();
                    beganUpdate = true;
                }
                catch (System.ComponentModel.Win32Exception)
                {
                }

                _auditIssueList.Items.Clear();
                for (int i = 0; i < _auditReviewIssues.Count; i++)
                {
                    AuditReviewIssue issue = _auditReviewIssues[i];
                    if (issue == null)
                        continue;

                    int index = _auditIssueList.Items.Add(issue);
                    if (issue.AutoApplicable && !issue.Applied)
                        _auditIssueList.SetItemChecked(index, true);
                }
            }
            catch (System.ComponentModel.Win32Exception)
            {
            }
            catch (InvalidOperationException)
            {
            }
            catch (ObjectDisposedException)
            {
            }
            finally
            {
                if (beganUpdate)
                {
                    try { _auditIssueList.EndUpdate(); } catch { }
                }
                _auditRenderingIssues = false;
            }

            try
            {
                if (selectFirst && _auditIssueList.Items.Count > 0)
                    _auditIssueList.SelectedIndex = 0;
                else if (_auditIssueList.Items.Count == 0)
                    ClearAuditPreview();
            }
            catch (System.ComponentModel.Win32Exception)
            {
            }
            catch (InvalidOperationException)
            {
            }
            catch (ObjectDisposedException)
            {
            }

            try { UpdateAuditActionButtons(); } catch { }
        }

'@
$panel = $panel.Substring(0, $renderStart) + $renderMethod.Replace("`n", $panelNl) + $panel.Substring($renderEnd)

$itemStart = $panel.IndexOf('        private void AuditIssueList_ItemCheck(object sender, ItemCheckEventArgs e)')
$itemEnd = if ($itemStart -ge 0) { $panel.IndexOf('        private void AuditIssueList_SelectedIndexChanged(', $itemStart) } else { -1 }
if ($itemStart -lt 0 -or $itemEnd -le $itemStart) {
    throw 'AuditIssueList_ItemCheck method was not found.'
}

$itemMethod = @'
        private void QueueAuditActionButtonsRefresh()
        {
            if (_auditIssueList == null || _auditIssueList.IsDisposed || _auditIssueList.Disposing)
                return;

            if (_auditIssueList.IsHandleCreated)
            {
                try
                {
                    _auditIssueList.BeginInvoke(new Action(() =>
                    {
                        try { UpdateAuditActionButtons(); } catch { }
                    }));
                    return;
                }
                catch (System.ComponentModel.Win32Exception)
                {
                }
                catch (InvalidOperationException)
                {
                }
            }

            try { UpdateAuditActionButtons(); } catch { }
        }

        private void AuditIssueList_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (_auditRenderingIssues ||
                _auditIssueList == null ||
                _auditIssueList.IsDisposed ||
                _auditIssueList.Disposing)
            {
                return;
            }

            if (e.Index < 0 || e.Index >= _auditIssueList.Items.Count)
                return;

            AuditReviewIssue issue = _auditIssueList.Items[e.Index] as AuditReviewIssue;
            if (issue != null && (!issue.AutoApplicable || issue.Applied) && e.NewValue != CheckState.Unchecked)
            {
                e.NewValue = CheckState.Unchecked;
                try
                {
                    if (_auditReviewHeader != null && !_auditReviewHeader.IsDisposed && !_auditReviewHeader.Disposing)
                    {
                        _auditReviewHeader.Text =
                            "\u042d\u0442\u043e \u0437\u0430\u043c\u0435\u0447\u0430\u043d\u0438\u0435 \u0442\u0440\u0435\u0431\u0443\u0435\u0442 \u0440\u0435\u0448\u0435\u043d\u0438\u044f \u0430\u0432\u0442\u043e\u0440\u0430 \u0438 \u043d\u0435 \u0431\u0443\u0434\u0435\u0442 \u043f\u0440\u0438\u043c\u0435\u043d\u0435\u043d\u043e \u0430\u0432\u0442\u043e\u043c\u0430\u0442\u0438\u0447\u0435\u0441\u043a\u0438.";
                    }
                }
                catch (System.ComponentModel.Win32Exception)
                {
                }
                catch (InvalidOperationException)
                {
                }
            }

            QueueAuditActionButtonsRefresh();
        }

'@
$panel = $panel.Substring(0, $itemStart) + $itemMethod.Replace("`n", $panelNl) + $panel.Substring($itemEnd)

# Keep the post-apply count null-safe even before/after other build-time patches.
$panel = $panel.Replace(
    '_auditReviewIssues.Count(i => i.AutoApplicable && !i.Applied)',
    '_auditReviewIssues.Count(i => i != null && i.AutoApplicable && !i.Applied)'
)

# Applying from the review panel has its own cleanup path. It is UI-only cleanup,
# so it must not escape after Word edits have already succeeded.
$applyStart = $panel.IndexOf('        private async void AuditApplySelectedButton_Click(object sender, EventArgs e)')
$applyEnd = if ($applyStart -ge 0) { $panel.IndexOf('        private void AuditSelectSafeButton_Click(', $applyStart) } else { -1 }
if ($applyStart -lt 0 -or $applyEnd -le $applyStart) {
    throw 'AuditApplySelectedButton_Click method was not found.'
}
$apply = $panel.Substring($applyStart, $applyEnd - $applyStart)
$applyFinallyPattern = '(?s)            finally\s*\{\s*_auditReviewBusy = false;\s*SetAuditReviewControlsEnabled\(true\);\s*\}'
$applyFinally = @'
            finally
            {
                _auditReviewBusy = false;
                try { SetAuditReviewControlsEnabled(true); } catch { }
            }
'@
if ([regex]::IsMatch($apply, $applyFinallyPattern)) {
    $apply = [regex]::Replace($apply, $applyFinallyPattern, $applyFinally.TrimEnd().Replace("`n", $panelNl), 1)
} elseif (-not $apply.Contains('try { SetAuditReviewControlsEnabled(true); } catch { }')) {
    throw 'Audit panel apply cleanup block was not found.'
}
$panel = $panel.Substring(0, $applyStart) + $apply + $panel.Substring($applyEnd)

if (-not $fix.Contains('TrySetAuditControlEnabled')) {
    throw 'AuditFix control handle guard is missing.'
}
if (-not $fix.Contains('_responseTextBox.IsHandleCreated')) {
    throw 'AuditFix response scroll handle guard is missing.'
}
if (-not $panel.Contains('private bool _auditRenderingIssues;')) {
    throw 'Audit render suppression flag is missing.'
}
if (-not $panel.Contains('QueueAuditActionButtonsRefresh')) {
    throw 'Audit item-check handle guard is missing.'
}

Write-Utf8Text $panelPath $panel
Write-Host 'Audit apply UI-handle stability patch applied successfully.'
