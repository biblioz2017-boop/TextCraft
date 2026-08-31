$ErrorActionPreference = 'Stop'

# Keep this script strictly ASCII for Windows PowerShell 5.1.
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$path = 'GenerateUserControl.AuditPanel.cs'
$fullPath = (Resolve-Path $path).Path
$text = [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8)
$nl = if ($text.Contains("`r`n")) { "`r`n" } else { "`n" }

# The post-apply diagnostic patch may be structurally replaced by the headless
# review patch. Keep headless diagnostics self-contained instead of depending on
# a helper introduced by an earlier build-time patch.
$text = $text.Replace('LogAuditUiException(', 'LogHeadlessAuditException(')

if (-not $text.Contains('private static void LogHeadlessAuditException(')) {
    $anchor = '        private void AppendAuditReviewSummaryToResponse()'
    $index = $text.IndexOf($anchor)
    if ($index -lt 0) {
        throw 'Headless audit summary method anchor was not found.'
    }

    $helper = @'
        private static void LogHeadlessAuditException(string stage, Exception ex)
        {
            try
            {
                string logPath = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(),
                    "NeZnaika-audit-errors.log"
                );

                var entry = new StringBuilder();
                entry.AppendLine(DateTime.Now.ToString("O"));
                entry.Append("Stage: ").AppendLine(stage ?? string.Empty);
                try
                {
                    entry.Append("Process handles: ")
                        .AppendLine(System.Diagnostics.Process.GetCurrentProcess().HandleCount.ToString());
                }
                catch
                {
                }
                entry.AppendLine(ex == null ? string.Empty : ex.ToString());
                entry.AppendLine(new string('-', 72));

                System.IO.File.AppendAllText(
                    logPath,
                    entry.ToString(),
                    new System.Text.UTF8Encoding(false)
                );
            }
            catch
            {
            }
        }

'@

    $text = $text.Substring(0, $index) + $helper.Replace("`n", $nl) + $text.Substring($index)
}

if ($text.Contains('LogAuditUiException(')) {
    throw 'Legacy audit logger calls remain after headless logger repair.'
}
if (-not $text.Contains('private static void LogHeadlessAuditException(')) {
    throw 'Headless audit logger helper was not inserted.'
}

[System.IO.File]::WriteAllText($fullPath, $text, $utf8NoBom)
Write-Host 'Headless audit logger repair applied successfully.'
