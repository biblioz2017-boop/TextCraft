$ErrorActionPreference = 'Stop'

# Keep this script ASCII-only for Windows PowerShell 5.1.
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Read-Utf8Text([string]$Path) {
    $fullPath = (Resolve-Path $Path).Path
    return [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8)
}

function Write-Utf8Text([string]$Path, [string]$Text) {
    $fullPath = (Resolve-Path $Path).Path
    [System.IO.File]::WriteAllText($fullPath, $Text, $utf8NoBom)
}

Write-Host 'Wiring scientific audit auto-fix and review controls...'
$path = 'GenerateUserControl.Science.cs'
$science = Read-Utf8Text $path

if (-not $science.Contains('InitializeAuditFixControls();')) {
    $anchor = '            AddScientificQuickActions();'
    if (-not $science.Contains($anchor)) {
        throw 'Could not locate AddScientificQuickActions call.'
    }
    $science = $science.Replace(
        $anchor,
        $anchor + "`r`n            InitializeAuditFixControls();"
    )
}

if (-not $science.Contains('InitializeAuditReviewPanel();')) {
    $anchor = '            InitializeAuditFixControls();'
    if (-not $science.Contains($anchor)) {
        throw 'Could not locate InitializeAuditFixControls call.'
    }
    $science = $science.Replace(
        $anchor,
        $anchor + "`r`n            InitializeAuditReviewPanel();"
    )
}

if (-not $science.Contains('PrepareAuditTargetForRequest(templateName, anchorRange, userQuery);')) {
    $anchor = '                string templateName = GetSelectedTemplateName();'
    if (-not $science.Contains($anchor)) {
        throw 'Could not locate RAG-aware template selection.'
    }
    $science = $science.Replace(
        $anchor,
        $anchor + "`r`n                PrepareAuditTargetForRequest(templateName, anchorRange, userQuery);"
    )
}

if (-not $science.Contains('CaptureAuditResultIfNeeded(templateName, response);')) {
    $anchor = "                _lastResponseMarkdown = response;`r`n                _lastTemplateName = templateName;"
    if (-not $science.Contains($anchor)) {
        throw 'Could not locate RAG-aware response capture block.'
    }
    $science = $science.Replace(
        $anchor,
        $anchor + "`r`n                CaptureAuditResultIfNeeded(templateName, response);"
    )
}

if (-not $science.Contains('await BuildAuditReviewPanelAsync();')) {
    $anchor = '                CaptureAuditResultIfNeeded(templateName, response);'
    if (-not $science.Contains($anchor)) {
        throw 'Could not locate audit result capture call.'
    }

    # Use C# Unicode escapes so this PowerShell file stays ASCII-only.
    $auditName = '\u041D\u0430\u0443\u0447\u043D\u044B\u0439 \u0430\u0443\u0434\u0438\u0442'
    $extra = "`r`n                if (string.Equals(templateName, `"$auditName`", StringComparison.Ordinal) && responseAccepted)`r`n                    await BuildAuditReviewPanelAsync();"
    $science = $science.Replace($anchor, $anchor + $extra)
}

if (-not $science.Contains('RememberAuditTarget(selection.Range.Duplicate, auditedText);')) {
    $old = @'
                int maxTokens = Math.Max(900, (int)(ThisAddIn.ContextLength * 0.42));
                PromptTextBox.Text = CommonUtils.SubstringTokens(
                    (selection.Text ?? string.Empty).Trim(),
                    maxTokens
                );
                GenerateButton.PerformClick();
'@

    $new = @'
                int maxTokens = Math.Max(900, (int)(ThisAddIn.ContextLength * 0.42));
                string auditedText = CommonUtils.SubstringTokens(
                    (selection.Text ?? string.Empty).Trim(),
                    maxTokens
                );
                RememberAuditTarget(selection.Range.Duplicate, auditedText);
                PromptTextBox.Text = auditedText;
                GenerateButton.PerformClick();
'@

    if (-not $science.Contains($old.Trim())) {
        throw 'Could not locate AuditChapterButton_Click text handoff.'
    }
    $science = $science.Replace($old.Trim(), $new.Trim())
}

if (-not $science.Contains('PrepareAuditReviewForNewRun();')) {
    $anchor = '                RememberAuditTarget(selection.Range.Duplicate, auditedText);'
    if (-not $science.Contains($anchor)) {
        throw 'Could not locate RememberAuditTarget call.'
    }
    $science = $science.Replace(
        $anchor,
        $anchor + "`r`n                PrepareAuditReviewForNewRun();"
    )
}

Write-Utf8Text $path $science
Write-Host 'Scientific audit auto-fix and review wiring applied.'
