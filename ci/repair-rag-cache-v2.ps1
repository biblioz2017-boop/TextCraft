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

Write-Host 'Applying RAG cache v2 validation...'

$path = 'RAGControl.Science.cs'
$text = Read-Utf8Text $path
$nl = Get-NewLine $text

if (-not $text.Contains('NeZnaika-RAG-cache-v2|')) {
    $oldIdentity = @'
            string identity =
                Path.GetFullPath(filePath) + "|" +
'@
    $newIdentity = @'
            string identity =
                "NeZnaika-RAG-cache-v2|" +
                Path.GetFullPath(filePath) + "|" +
'@
    $oldIdentity = $oldIdentity.TrimEnd("`r", "`n").Replace("`n", $nl)
    $newIdentity = $newIdentity.TrimEnd("`r", "`n").Replace("`n", $nl)
    if (-not $text.Contains($oldIdentity)) {
        throw 'Could not locate persistent RAG cache identity.'
    }
    $text = $text.Replace($oldIdentity, $newIdentity)
}

if (-not $text.Contains('cached.Indexes.Values.Any(index => index != null && index.Count > 0)')) {
    $oldLoad = @'
                var cached = new HyperVectorDB.HyperVectorDB(ThisAddIn.Embedder, cachePath);
                cached.Load();
                database = cached;
                return true;
'@
    $newLoad = @'
                var cached = new HyperVectorDB.HyperVectorDB(ThisAddIn.Embedder, cachePath);
                cached.Load();

                bool hasDocuments = cached.Indexes != null &&
                    cached.Indexes.Values.Any(index => index != null && index.Count > 0);
                if (!hasDocuments)
                    return false;

                database = cached;
                return true;
'@
    $oldLoad = $oldLoad.TrimEnd("`r", "`n").Replace("`n", $nl)
    $newLoad = $newLoad.TrimEnd("`r", "`n").Replace("`n", $nl)
    if (-not $text.Contains($oldLoad)) {
        throw 'Could not locate persistent RAG cache load block.'
    }
    $text = $text.Replace($oldLoad, $newLoad)
}

Write-Utf8Text $path $text

$verify = Read-Utf8Text $path
if (-not $verify.Contains('NeZnaika-RAG-cache-v2|')) {
    throw 'RAG cache v2 identity is missing.'
}
if (-not $verify.Contains('cached.Indexes.Values.Any(index => index != null && index.Count > 0)')) {
    throw 'RAG cache empty-index validation is missing.'
}

Write-Host 'RAG cache v2 validation applied successfully.'
