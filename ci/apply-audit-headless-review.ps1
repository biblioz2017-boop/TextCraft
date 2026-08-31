$ErrorActionPreference = 'Stop'

# Keep this script strictly ASCII for Windows PowerShell 5.1.
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

function Replace-Method(
    [string]$Text,
    [string]$StartMarker,
    [string]$EndMarker,
    [string]$Replacement,
    [string]$NewLine
) {
    $start = $Text.IndexOf($StartMarker)
    if ($start -lt 0) { throw ('Start marker not found: ' + $StartMarker) }
    $end = $Text.IndexOf($EndMarker, $start)
    if ($end -le $start) { throw ('End marker not found: ' + $EndMarker) }
    return $Text.Substring(0, $start) + $Replacement.Replace("`n", $NewLine) + $Text.Substring($end)
}

Write-Host 'Applying headless audit review mode...'

$path = 'GenerateUserControl.AuditPanel.cs'
$panel = Read-Utf8Text $path
$nl = Get-NewLine $panel

# Do not create the extra WinForms panel. The ListBox/CheckedListBox was the
# source of NativeUpdateSelection, Graphics.FromHwndInternal and handle errors.
$initialize = @'
        private void InitializeAuditReviewPanel()
        {
            // Headless review mode: stage 2 still runs through the LLM, but no
            // secondary WinForms review panel or CheckedListBox is created.
            _auditReviewPanel = null;
            _auditReviewLayout = null;
            _auditIssueList = null;
            _auditReviewActions = null;
        }

'@
$panel = Replace-Method $panel `
    '        private void InitializeAuditReviewPanel()' `
    '        private static RichTextBox CreateAuditPreviewBox()' `
    $initialize `
    $nl

# Starting a new audit must not touch any review-panel controls.
$prepare = @'
        private void PrepareAuditReviewForNewRun()
        {
            _auditReviewIssues.Clear();
            _auditReviewProgressStartedUtc = DateTime.UtcNow;
            _auditReviewStreamedCharacters = 0;
            _auditReviewProgressPhase = string.Empty;

            if (_evidencePanel != null)
                _evidencePanel.Visible = true;
        }

'@
$panel = Replace-Method $panel `
    '        private void PrepareAuditReviewForNewRun()' `
    '        private void StartAuditReviewProgress()' `
    $prepare `
    $nl

# Applying audit fixes should update only the in-memory issue state. Repainting
# the removed ListBox after Word edits is both unnecessary and unsafe.
$mark = @'
        private void MarkAuditReviewEditsApplied(IEnumerable<AuditEdit> edits)
        {
            if (edits == null)
                return;

            var appliedKeys = new HashSet<string>(StringComparer.Ordinal);
            foreach (AuditEdit edit in edits)
            {
                if (edit != null)
                {
                    appliedKeys.Add(
                        (edit.FindText ?? string.Empty) + "\n" +
                        (edit.Replacement ?? string.Empty)
                    );
                }
            }

            foreach (AuditReviewIssue issue in _auditReviewIssues)
            {
                if (issue == null)
                    continue;

                string key =
                    (issue.FindText ?? string.Empty) + "\n" +
                    (issue.Replacement ?? string.Empty);
                if (appliedKeys.Contains(key))
                    issue.Applied = true;
            }
        }

'@
$panel = Replace-Method $panel `
    '        private void MarkAuditReviewEditsApplied(IEnumerable<AuditEdit> edits)' `
    '        private async Task BuildAuditReviewPanelAsync()' `
    $mark `
    $nl

# Stage 2 remains mandatory LLM processing. Its parsed result is stored in
# _auditReviewIssues and rendered as plain text in the already-existing response
# box. No new window handles are created.
$build = @'
        private async Task BuildAuditReviewPanelAsync()
        {
            if (_auditReviewBusy || !HasPendingAuditReview())
                return;

            Word.Document document;
            Word.Range targetRange;
            if (!TryGetCurrentAuditTarget(out document, out targetRange))
                return;

            string currentText = targetRange.Text ?? string.Empty;
            if (string.IsNullOrWhiteSpace(currentText))
                return;

            try
            {
                _auditReviewBusy = true;
                SetAuditFixButtons(false);
                Forge.SetModelActivity(true, "\u0421\u0442\u0440\u0443\u043a\u0442\u0443\u0440\u0438\u0440\u0443\u0435\u0442 \u0430\u0443\u0434\u0438\u0442\u2026");
                StartAuditReviewProgress();

                if (_responseLabel != null && !_responseLabel.IsDisposed)
                    _responseLabel.Text = "\u0410\u0443\u0434\u0438\u0442 \u2014 \u044d\u0442\u0430\u043f 2 \u0438\u0437 2: LLM \u0441\u0442\u0440\u0443\u043a\u0442\u0443\u0440\u0438\u0440\u0443\u0435\u0442 \u0437\u0430\u043c\u0435\u0447\u0430\u043d\u0438\u044f\u2026";

                List<AuditReviewIssue> issues = await GenerateAuditReviewIssuesAsync(
                    currentText,
                    _lastAuditReport,
                    20
                );

                _auditReviewIssues.Clear();
                if (issues != null)
                    _auditReviewIssues.AddRange(issues.Where(i => i != null));

                AppendAuditReviewSummaryToResponse();
                CompleteAuditReviewProgress(_auditReviewIssues.Count);
            }
            catch (OperationCanceledException)
            {
                AppendAuditReviewStatusToResponse(
                    "\u0421\u0442\u0440\u0443\u043a\u0442\u0443\u0440\u0438\u0440\u043e\u0432\u0430\u043d\u0438\u0435 \u0430\u0443\u0434\u0438\u0442\u0430 \u043e\u0441\u0442\u0430\u043d\u043e\u0432\u043b\u0435\u043d\u043e \u043f\u043e\u043b\u044c\u0437\u043e\u0432\u0430\u0442\u0435\u043b\u0435\u043c."
                );
            }
            catch (Exception ex)
            {
                AppendAuditReviewStatusToResponse(
                    "\u041e\u0448\u0438\u0431\u043a\u0430 \u0432\u0442\u043e\u0440\u043e\u0433\u043e \u044d\u0442\u0430\u043f\u0430: " + ex.Message
                );
                try { LogAuditUiException("HeadlessBuildAuditReview", ex); } catch { }
            }
            finally
            {
                try { Forge.SetModelActivity(false, null); } catch { }
                _auditReviewBusy = false;
                try { SetAuditFixButtons(HasPendingAuditReview()); } catch { }

                try
                {
                    if (_responseLabel != null && !_responseLabel.IsDisposed)
                        _responseLabel.Text = "\u0414\u0438\u0430\u043b\u043e\u0433 / \u043e\u0442\u0432\u0435\u0442:";
                }
                catch
                {
                }
            }
        }

        private void AppendAuditReviewSummaryToResponse()
        {
            if (_responseTextBox == null ||
                _responseTextBox.IsDisposed ||
                _responseTextBox.Disposing ||
                !_responseTextBox.IsHandleCreated)
            {
                return;
            }

            var summary = new StringBuilder();
            summary.AppendLine();
            summary.AppendLine();
            summary.AppendLine("[\u041d\u0435Z\u043d\u0430\u0439\u043a\u0430 \u2014 \u0441\u0442\u0440\u0443\u043a\u0442\u0443\u0440\u0438\u0440\u043e\u0432\u0430\u043d\u043d\u044b\u0435 \u0437\u0430\u043c\u0435\u0447\u0430\u043d\u0438\u044f]");

            if (_auditReviewIssues.Count == 0)
            {
                summary.AppendLine("LLM \u043d\u0435 \u0432\u0435\u0440\u043d\u0443\u043b\u0430 \u0437\u0430\u043c\u0435\u0447\u0430\u043d\u0438\u0439, \u043a\u043e\u0442\u043e\u0440\u044b\u0435 \u043c\u043e\u0436\u043d\u043e \u043e\u0434\u043d\u043e\u0437\u043d\u0430\u0447\u043d\u043e \u043f\u0440\u0438\u0432\u044f\u0437\u0430\u0442\u044c \u043a \u0442\u0435\u043a\u0441\u0442\u0443.");
            }
            else
            {
                int number = 1;
                foreach (AuditReviewIssue issue in _auditReviewIssues)
                {
                    if (issue == null)
                        continue;

                    summary.Append(number++).Append(". [")
                        .Append(string.IsNullOrWhiteSpace(issue.Category) ? "\u0417\u0430\u043c\u0435\u0447\u0430\u043d\u0438\u0435" : issue.Category.Trim())
                        .Append("] ")
                        .AppendLine(issue.Reason ?? string.Empty);

                    if (!string.IsNullOrWhiteSpace(issue.FindText))
                        summary.Append("   \u0411\u044b\u043b\u043e: ").AppendLine(issue.FindText.Trim());

                    if (issue.AutoApplicable && !string.IsNullOrWhiteSpace(issue.Replacement))
                        summary.Append("   \u0421\u0442\u0430\u043b\u043e: ").AppendLine(issue.Replacement.Trim());
                    else
                        summary.AppendLine("   \u0420\u0435\u0436\u0438\u043c: \u0440\u0443\u0447\u043d\u0430\u044f \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0430.");
                }
            }

            int safeCount = _auditReviewIssues.Count(
                i => i != null && i.AutoApplicable && !i.Applied
            );
            summary.AppendLine();
            summary.Append("\u0411\u0435\u0437\u043e\u043f\u0430\u0441\u043d\u044b\u0445 \u043f\u0440\u0430\u0432\u043e\u043a \u0434\u043b\u044f \u043a\u043d\u043e\u043f\u043a\u0438 \u00ab\u0418\u0441\u043f\u0440\u0430\u0432\u0438\u0442\u044c \u0430\u0443\u0434\u0438\u0442\u00bb: ")
                .Append(safeCount);

            try
            {
                _responseTextBox.AppendText(summary.ToString());
            }
            catch (System.ComponentModel.Win32Exception ex)
            {
                try { LogAuditUiException("AppendAuditReviewSummary", ex); } catch { }
            }
            catch (InvalidOperationException ex)
            {
                try { LogAuditUiException("AppendAuditReviewSummary", ex); } catch { }
            }
        }

        private void AppendAuditReviewStatusToResponse(string text)
        {
            if (_responseTextBox == null ||
                _responseTextBox.IsDisposed ||
                _responseTextBox.Disposing ||
                !_responseTextBox.IsHandleCreated)
            {
                return;
            }

            try
            {
                _responseTextBox.AppendText(
                    Environment.NewLine + Environment.NewLine +
                    "[\u041d\u0435Z\u043d\u0430\u0439\u043a\u0430 \u2014 \u0430\u0443\u0434\u0438\u0442] " +
                    (text ?? string.Empty)
                );
            }
            catch
            {
            }
        }

'@
$panel = Replace-Method $panel `
    '        private async Task BuildAuditReviewPanelAsync()' `
    '        private async Task<List<AuditReviewIssue>> GenerateAuditReviewIssuesWithFallbackAsync(' `
    $build `
    $nl

# If the fallback wrapper is absent in a future source shape, retry against the
# direct generator anchor.
if (-not $panel.Contains('private void AppendAuditReviewSummaryToResponse()')) {
    throw 'Headless audit review helper was not inserted.'
}

# No review-panel control state should be queried after this final patch.
$controlsStart = $panel.IndexOf('        private void SetAuditReviewControlsEnabled(bool enabled)')
$controlsEnd = if ($controlsStart -ge 0) {
    $panel.IndexOf('        private void ShowAuditReviewPanel()', $controlsStart)
} else { -1 }
if ($controlsStart -ge 0 -and $controlsEnd -gt $controlsStart) {
    $noopControls = @'
        private void SetAuditReviewControlsEnabled(bool enabled)
        {
            // Headless mode intentionally has no review-panel controls.
        }

'@
    $panel = $panel.Substring(0, $controlsStart) + $noopControls.Replace("`n", $nl) + $panel.Substring($controlsEnd)
}

$showStart = $panel.IndexOf('        private void ShowAuditReviewPanel()')
$showEnd = if ($showStart -ge 0) { $panel.IndexOf('        private void HideAuditReviewPanel()', $showStart) } else { -1 }
if ($showStart -ge 0 -and $showEnd -gt $showStart) {
    $noopShow = @'
        private void ShowAuditReviewPanel()
        {
            // Removed: unstable secondary WinForms review panel.
        }

'@
    $panel = $panel.Substring(0, $showStart) + $noopShow.Replace("`n", $nl) + $panel.Substring($showEnd)
}

$hideStart = $panel.IndexOf('        private void HideAuditReviewPanel()')
$hideEnd = if ($hideStart -ge 0) { $panel.IndexOf('        private void ResetAuditReviewPanelState(', $hideStart) } else { -1 }
if ($hideStart -ge 0 -and $hideEnd -gt $hideStart) {
    $noopHide = @'
        private void HideAuditReviewPanel()
        {
            if (_evidencePanel != null)
                _evidencePanel.Visible = true;
        }

'@
    $panel = $panel.Substring(0, $hideStart) + $noopHide.Replace("`n", $nl) + $panel.Substring($hideEnd)
}

if (-not $panel.Contains('Headless review mode: stage 2 still runs through the LLM')) {
    throw 'Headless review initialization marker is missing.'
}
if (-not $panel.Contains('List<AuditReviewIssue> issues = await GenerateAuditReviewIssuesAsync(')) {
    throw 'Mandatory LLM stage-2 call is missing from headless review mode.'
}
if (-not $panel.Contains('private void AppendAuditReviewSummaryToResponse()')) {
    throw 'Headless structured audit text output is missing.'
}

Write-Utf8Text $path $panel
Write-Host 'Headless audit review mode applied successfully.'
