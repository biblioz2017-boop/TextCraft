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

Write-Host 'Applying final chat scroll-range fix...'

$chatPath = 'GenerateUserControl.cs'
$chat = Read-Utf8Text $chatPath
$nl = Get-NewLine $chat

# Make the response box explicitly vertical-scrollable. WordWrap prevents a
# horizontal range from interfering with the native RichEdit scroll metrics.
if (-not $chat.Contains('ScrollBars = RichTextBoxScrollBars.Vertical')) {
    $oldBox =
        '                HideSelection = false,' + $nl +
        '                BackColor = SystemColors.Window,'
    $newBox =
        '                HideSelection = false,' + $nl +
        '                ScrollBars = RichTextBoxScrollBars.Vertical,' + $nl +
        '                WordWrap = true,' + $nl +
        '                BackColor = SystemColors.Window,'

    if (-not $chat.Contains($oldBox)) {
        throw 'Could not locate response RichTextBox initializer.'
    }
    $chat = $chat.Replace($oldBox, $newBox)
}

# Centralize the final native RichEdit scroll operation. Calling this after
# WM_SETREDRAW is re-enabled forces the control to recalculate its real bottom.
if (-not $chat.Contains('private void ScrollResponseToEnd()')) {
    $anchor = '        private void AppendConversationHeader(string question, string templateName)'
    if (-not $chat.Contains($anchor)) {
        throw 'Could not locate chat header method for scroll helper insertion.'
    }

    $helper = @'
        private void ScrollResponseToEnd()
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
                _responseTextBox.SelectionStart = _responseTextBox.TextLength;
                _responseTextBox.SelectionLength = 0;
                _responseTextBox.ScrollToCaret();
                _responseTextBox.Invalidate();
                _responseTextBox.Update();
            }
            catch
            {
            }
        }

'@
    $chat = $chat.Replace($anchor, $helper.Replace("`n", $nl) + $anchor)
}

# Use the same helper during ordinary streaming as well.
$oldScroll =
    '            _responseTextBox.SelectionStart = _responseTextBox.TextLength;' + $nl +
    '            _responseTextBox.ScrollToCaret();'
if ($chat.Contains($oldScroll)) {
    $chat = $chat.Replace($oldScroll, '            ScrollResponseToEnd();')
}

$oldStreamScroll =
    '                            _responseTextBox.SelectionStart = _responseTextBox.TextLength;' + $nl +
    '                            _responseTextBox.ScrollToCaret();'
if ($chat.Contains($oldStreamScroll)) {
    $chat = $chat.Replace($oldStreamScroll, '                            ScrollResponseToEnd();')
}

Write-Utf8Text $chatPath $chat

$sciencePath = 'GenerateUserControl.Science.cs'
$science = Read-Utf8Text $sciencePath
$snl = Get-NewLine $science

# The smooth-streaming hook temporarily disables native redraw. Refresh alone is
# not enough: RichEdit can retain an obsolete vertical scroll range. Re-enable
# redraw and immediately move the caret to the true end.
$oldTimerFinish =
    '                    StopChatRedrawTimer();' + $snl +
    '                    SetChatRedraw(true);' + $snl +
    '                    _responseTextBox.Refresh();' + $snl +
    '                    return;'
$newTimerFinish =
    '                    StopChatRedrawTimer();' + $snl +
    '                    SetChatRedraw(true);' + $snl +
    '                    ScrollResponseToEnd();' + $snl +
    '                    _responseTextBox.Refresh();' + $snl +
    '                    return;'
if ($science.Contains($oldTimerFinish)) {
    $science = $science.Replace($oldTimerFinish, $newTimerFinish)
} elseif (-not $science.Contains('                    ScrollResponseToEnd();' + $snl + '                    _responseTextBox.Refresh();')) {
    throw 'Could not locate smooth-streaming completion branch.'
}

# Do not wait for the timer tick at the end of a request. This is also the final
# path after the second audit stage, so structured audit text becomes reachable.
$finallyAnchor =
    '                GenerateButton.Enabled = true;' + $snl +
    '                if (_responseLabel != null)'
$finallyInsert =
    '                GenerateButton.Enabled = true;' + $snl +
    '                StopChatRedrawTimer();' + $snl +
    '                SetChatRedraw(true);' + $snl +
    '                ScrollResponseToEnd();' + $snl +
    '                if (_responseLabel != null)'
if ($science.Contains($finallyAnchor)) {
    $science = $science.Replace($finallyAnchor, $finallyInsert)
} elseif (-not $science.Contains('                StopChatRedrawTimer();' + $snl + '                SetChatRedraw(true);' + $snl + '                ScrollResponseToEnd();')) {
    throw 'Could not locate RAG-aware finalization block.'
}

Write-Utf8Text $sciencePath $science

$panelPath = 'GenerateUserControl.AuditPanel.cs'
if (Test-Path $panelPath) {
    $panel = Read-Utf8Text $panelPath
    $pnl = Get-NewLine $panel

    $summaryAppend = '                _responseTextBox.AppendText(summary.ToString());'
    if ($panel.Contains($summaryAppend) -and
        -not $panel.Contains($summaryAppend + $pnl + '                ScrollResponseToEnd();')) {
        $panel = $panel.Replace(
            $summaryAppend,
            $summaryAppend + $pnl + '                ScrollResponseToEnd();'
        )
    }

    Write-Utf8Text $panelPath $panel
}

if (-not (Read-Utf8Text $chatPath).Contains('private void ScrollResponseToEnd()')) {
    throw 'Final chat scroll helper is missing.'
}
if (-not (Read-Utf8Text $sciencePath).Contains('ScrollResponseToEnd();')) {
    throw 'Final chat scroll call is missing from science UI.'
}

Write-Host 'Final chat scroll-range fix applied successfully.'
