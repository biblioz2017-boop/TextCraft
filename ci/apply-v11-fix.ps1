$ErrorActionPreference = 'Stop'

$path = 'ci/build-patched.ps1'
$scriptText = Get-Content $path -Raw

function Replace-Required {
    param(
        [string]$Text,
        [string]$Old,
        [string]$New,
        [string]$Label
    )
    if (-not $Text.Contains($Old)) {
        throw "Required patch target not found: $Label"
    }
    return $Text.Replace($Old, $New)
}

$scriptText = Replace-Required $scriptText `
    '            IEnumerable<string> fileContent;' `
    '            List<string> fileContent;' `
    'fileContent type'

$oldRead = @'
                fileContent = await ReadPdfFileAsync(filePath, CHUNK_LEN);
'@
$newRead = @'
                fileContent = (await ReadPdfFileAsync(filePath, CHUNK_LEN)).ToList();
                if (fileContent.Count == 0)
                {
                    MarkFileStatus(filePath, $"[NO TEXT] {Path.GetFileName(filePath)} (0 chunks)");
                    _activeIndexVersions.TryRemove(filePath, out _);
                    return;
                }
                MarkFileStatus(filePath, $"[INDEXING] {Path.GetFileName(filePath)} ({fileContent.Count} chunks)");
'@
$scriptText = Replace-Required $scriptText $oldRead $newRead 'PDF read/status block'

$oldDb = @'
            var stagedDb = new HyperVectorDB.HyperVectorDB(ThisAddIn.Embedder, Path.GetTempPath());
'@
$newDb = @'
            var stagedDb = new HyperVectorDB.HyperVectorDB(ThisAddIn.Embedder, Path.GetTempPath());
            if (!stagedDb.CreateIndex(filePath))
                throw new InvalidOperationException($"Could not create vector index for {filePath}");
'@
$scriptText = Replace-Required $scriptText $oldDb $newDb 'named HyperVectorDB index creation'

$scriptText = Replace-Required $scriptText `
    '                    int fileContentCount = fileContent.Count();' `
    '                    int fileContentCount = fileContent.Count;' `
    'chunk count'

$oldIndex = @'
                        stagedDb.IndexDocument(fileContent.ElementAt(i));
'@
$newIndex = @'
                        if (!stagedDb.IndexDocument(filePath, fileContent[i]))
                            throw new InvalidOperationException($"Vector indexing failed for {Path.GetFileName(filePath)}, chunk {i + 1}/{fileContentCount}");
'@
$scriptText = Replace-Required $scriptText $oldIndex $newIndex 'checked named chunk insertion'

$oldPublish = @'
            long currentVersion;
            if (_activeIndexVersions.TryGetValue(filePath, out currentVersion) && currentVersion == indexVersion)
                _fileDatabases[filePath] = stagedDb;
        }

        private void RemoveFileEntry(string filePath)
'@
$newPublish = @'
            long currentVersion;
            if (_activeIndexVersions.TryGetValue(filePath, out currentVersion) && currentVersion == indexVersion)
            {
                _fileDatabases[filePath] = stagedDb;
                MarkFileStatus(filePath, $"[OK] {Path.GetFileName(filePath)} ({fileContent.Count} chunks)");
            }
        }

        private void MarkFileStatus(string filePath, string label)
        {
            Action update = () =>
            {
                for (int i = 0; i < _fileList.Count; i++)
                {
                    if (string.Equals(_fileList[i].Value, filePath, StringComparison.OrdinalIgnoreCase))
                    {
                        _fileList[i] = new KeyValuePair<string, string>(label, filePath);
                        break;
                    }
                }
            };

            if (this.InvokeRequired)
                this.Invoke((MethodInvoker)delegate { update(); });
            else
                update();
        }

        private void RemoveFileEntry(string filePath)
'@
$scriptText = Replace-Required $scriptText $oldPublish $newPublish 'publish/status method'

$oldEmpty = @'
            var databases = _fileDatabases.ToArray();
            if (databases.Length == 0 || maxTokens <= 0)
                return string.Empty;
'@
$newEmpty = @'
            var databases = _fileDatabases.ToArray();
            if (maxTokens <= 0)
                return string.Empty;
            if (databases.Length == 0)
            {
                if (_fileList != null && _fileList.Count > 0)
                    return "[RAG STATUS: files are attached but no PDF has completed vector indexing. Wait for [OK].]";
                return string.Empty;
            }
'@
$scriptText = Replace-Required $scriptText $oldEmpty $newEmpty 'empty RAG status'

$oldBuilder = @'
            StringBuilder ragContext = new StringBuilder();
'@
$newBuilder = @'
            StringBuilder ragContext = new StringBuilder();
            ragContext.AppendLine("[RAG SOURCES READY]");
            foreach (var source in databases.OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase))
                ragContext.AppendLine($"- {Path.GetFileName(source.Key)}");
            ragContext.AppendLine();
'@
$scriptText = Replace-Required $scriptText $oldBuilder $newBuilder 'RAG ready-source header'

$scriptText = Replace-Required $scriptText `
    "Write-Host 'Build preparation complete.'" `
    "Write-Host 'Build preparation complete.'`r`nexit 0" `
    'successful VSTO exit'

Set-Content $path $scriptText -Encoding UTF8
Write-Host 'TextCraft 1.0.11 RAG patch prepared.'
