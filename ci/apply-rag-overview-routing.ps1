$ErrorActionPreference = 'Stop'

# Keep this script strictly ASCII for Windows PowerShell 5.1.
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
function Read-Utf8Text([string]$Path) { return [System.IO.File]::ReadAllText((Resolve-Path $Path).Path, [System.Text.Encoding]::UTF8) }
function Write-Utf8Text([string]$Path, [string]$Text) { [System.IO.File]::WriteAllText((Resolve-Path $Path).Path, $Text, $utf8NoBom) }

$path = 'GenerateUserControl.Science.cs'
$text = Read-Utf8Text $path

$old = '                    forcedEvidence = await Task.Run(() => rag.GetRAGEvidence(retrievalQuery, 4));'
$new = @'
                    bool genericAttachedSourceRequest = IsGenericAttachedSourceRequest(userQuery);
                    if (genericAttachedSourceRequest)
                    {
                        if (_responseLabel != null)
                            _responseLabel.Text = "\u0414\u0438\u0430\u043b\u043e\u0433 / \u043e\u0442\u0432\u0435\u0442 \u2014 \u0441\u043e\u0431\u0438\u0440\u0430\u044e \u043e\u0431\u0437\u043e\u0440 \u043e\u0442\u043c\u0435\u0447\u0435\u043d\u043d\u044b\u0445 PDF\u2026";
                        forcedEvidence = await Task.Run(() => rag.GetCheckedPdfOverviewEvidence(10));
                    }
                    else
                    {
                        forcedEvidence = await Task.Run(() => rag.GetRAGEvidence(retrievalQuery, 4));
                    }
'@

if ($text.Contains($old)) {
    $text = $text.Replace($old, $new.TrimEnd("`r", "`n"))
} elseif (-not $text.Contains('GetCheckedPdfOverviewEvidence(10)')) {
    throw 'Could not locate strict-RAG evidence retrieval call for overview routing.'
}

Write-Utf8Text $path $text
if (-not (Read-Utf8Text $path).Contains('GetCheckedPdfOverviewEvidence(10)')) {
    throw 'Overview evidence routing was not applied.'
}
Write-Host 'Generic attached-PDF requests now use balanced overview evidence.'
