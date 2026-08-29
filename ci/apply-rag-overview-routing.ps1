$ErrorActionPreference = 'Stop'
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
                            _responseLabel.Text = "Диалог / ответ — собираю обзор отмеченных PDF…";
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

$oldNoEvidence = '                            "Строгий RAG остановил генерацию: в отмеченных PDF не найдено подходящих фрагментов по теме запроса. " +'
$newNoEvidence = @'
                            (genericAttachedSourceRequest
                                ? "Строгий RAG остановил генерацию: не удалось извлечь обзорные фрагменты из отмеченных PDF. "
                                : "Строгий RAG остановил генерацию: в отмеченных PDF не найдено подходящих фрагментов по теме запроса. ") +
'@
if ($text.Contains($oldNoEvidence)) {
    $text = $text.Replace($oldNoEvidence, $newNoEvidence.TrimEnd("`r", "`n"))
}

Write-Utf8Text $path $text
if (-not (Read-Utf8Text $path).Contains('GetCheckedPdfOverviewEvidence(10)')) {
    throw 'Overview evidence routing was not applied.'
}
Write-Host 'Generic attached-PDF requests now use balanced overview evidence.'
