$ErrorActionPreference = 'Stop'

Write-Host 'Patching RAGControl.cs...'
$path = 'RAGControl.cs'
$c = Get-Content $path -Raw
$c = $c.Replace('using System.Text;', "using System.Text;`r`nusing System.Text.RegularExpressions;`r`nusing System.Threading;")
$c = $c.Replace('public static readonly int CHUNK_LEN = CommonUtils.TokensToCharCount(256);', "public static readonly int CHUNK_LEN = CommonUtils.TokensToCharCount(640);`r`n        public static readonly int CHUNK_OVERLAP = CommonUtils.TokensToCharCount(96);")
$c = $c.Replace('int progressBarIncrement = (int)(fileContentCount * 0.1);', 'int progressBarIncrement = Math.Max(1, (int)Math.Ceiling(fileContentCount * 0.1));')

# Multi-document RAG: keep one independent HyperVectorDB per PDF. HyperVectorDB 1.0.6
# performs its all-index query with an unsafe shared List<T>, so querying several indexes
# inside one database can lose results or fail. A per-file database also lets us reserve
# RAG context for every attached PDF instead of allowing one file to dominate all hits.
$old = @'
        private ToolTip _fileToolTip = new ToolTip();
        private Queue<string> _removalQueue = new Queue<string>();
        private ConcurrentDictionary<int, int> _indexFileCount = new ConcurrentDictionary<int, int>();
        private BindingList<KeyValuePair<string, string>> _fileList; // Use KeyValuePair for label and filename
        private HyperVectorDB.HyperVectorDB _db;
        private bool _isIndexing;
        private float preciseProgressBar = 0;
'@
$new = @'
        private ToolTip _fileToolTip = new ToolTip();
        private ConcurrentDictionary<int, int> _indexFileCount = new ConcurrentDictionary<int, int>();
        private BindingList<KeyValuePair<string, string>> _fileList; // Use KeyValuePair for label and filename
        private ConcurrentDictionary<string, HyperVectorDB.HyperVectorDB> _fileDatabases = new ConcurrentDictionary<string, HyperVectorDB.HyperVectorDB>(StringComparer.OrdinalIgnoreCase);
        private ConcurrentDictionary<string, long> _activeIndexVersions = new ConcurrentDictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        private long _nextIndexVersion = 0;
        private float preciseProgressBar = 0;
'@
if (-not $c.Contains($old)) { throw 'Could not locate RAG database fields' }
$c = $c.Replace($old, $new)

$old = @'
            lock (Forge.InitializeDoor)
            {
                _db = new HyperVectorDB.HyperVectorDB(ThisAddIn.Embedder, Path.GetTempPath());
            }
'@
$new = @'
            _fileDatabases.Clear();
            _activeIndexVersions.Clear();
'@
if (-not $c.Contains($old)) { throw 'Could not locate RAG database initialization' }
$c = $c.Replace($old, $new)

$old = @'
        private void RemoveSelectedDocument()
        {
            string selectedDocument = FileListBox.SelectedItem.ToString();
            if (_isIndexing)
            {
                if (!_removalQueue.Contains(selectedDocument))
                    _removalQueue.Enqueue(selectedDocument);
            }
            else
            {
                DeleteDocument(selectedDocument);
            }
            _fileList.RemoveAt(FileListBox.SelectedIndex);
            AutoHideRemoveButton();
        }
'@
$new = @'
        private void RemoveSelectedDocument()
        {
            if (FileListBox.SelectedIndex < 0 || FileListBox.SelectedItem == null)
                return;

            var selectedItem = (KeyValuePair<string, string>)FileListBox.SelectedItem;
            string selectedDocument = selectedItem.Value;

            // Invalidates any in-flight indexing for this exact file. If the user removes
            // and immediately re-adds it, the new indexing receives a newer version and
            // the old task cannot publish stale vectors.
            _activeIndexVersions.TryRemove(selectedDocument, out _);
            DeleteDocument(selectedDocument);

            _fileList.RemoveAt(FileListBox.SelectedIndex);
            AutoHideRemoveButton();
        }
'@
if (-not $c.Contains($old)) { throw 'Could not locate RAG removal method' }
$c = $c.Replace($old, $new)

$startMulti = $c.IndexOf('        private async Task IndexDocumentAsync(string filePath)')
$endMulti = $c.IndexOf('        private void AutoHideRemoveButton()', $startMulti)
if ($startMulti -lt 0 -or $endMulti -lt 0) { throw 'Could not locate RAG indexing methods' }
$multiReplacement = @'
        private async Task IndexDocumentAsync(string filePath)
        {
            long indexVersion = Interlocked.Increment(ref _nextIndexVersion);
            _activeIndexVersions[filePath] = indexVersion;

            IEnumerable<string> fileContent;
            try
            {
                fileContent = await ReadPdfFileAsync(filePath, CHUNK_LEN);
            }
            catch
            {
                _activeIndexVersions.TryRemove(filePath, out _);
                RemoveFileEntry(filePath);
                throw;
            }

            // Build a complete private database first. It is published atomically only
            // after all chunks have embeddings, so Generate can safely keep using the
            // already-indexed PDFs while a newly added PDF is still processing.
            var stagedDb = new HyperVectorDB.HyperVectorDB(ThisAddIn.Embedder, Path.GetTempPath());

            try
            {
                await Task.Run(() =>
                {
                    int fileContentCount = fileContent.Count();
                    int progressBarIncrement = Math.Max(1, (int)Math.Ceiling(fileContentCount * 0.1));

                    for (int i = 0; i < fileContentCount; i++)
                    {
                        stagedDb.IndexDocument(fileContent.ElementAt(i));
                        if (i % progressBarIncrement == 0)
                        {
                            this.Invoke((MethodInvoker)delegate
                            {
                                UpdateProgressBar((float)progressBarIncrement / Math.Max(1, fileContentCount));
                            });
                        }
                    }
                });
            }
            catch
            {
                long ignoredVersion;
                _activeIndexVersions.TryRemove(filePath, out ignoredVersion);
                RemoveFileEntry(filePath);
                throw;
            }

            long currentVersion;
            if (_activeIndexVersions.TryGetValue(filePath, out currentVersion) && currentVersion == indexVersion)
                _fileDatabases[filePath] = stagedDb;
        }

        private void RemoveFileEntry(string filePath)
        {
            Action remove = () =>
            {
                var fileEntry = _fileList.FirstOrDefault(file => file.Value == filePath);
                if (fileEntry.Key != null)
                    _fileList.Remove(fileEntry);
                AutoHideRemoveButton();
            };

            if (this.InvokeRequired)
                this.Invoke((MethodInvoker)delegate { remove(); });
            else
                remove();
        }

        private bool DeleteDocument(string filePath)
        {
            HyperVectorDB.HyperVectorDB removed;
            return _fileDatabases.TryRemove(filePath, out removed);
        }

'@
$c = $c.Substring(0, $startMulti) + $multiReplacement + $c.Substring($endMulti)

$removeQueueStart = $c.IndexOf('        private void ProcessRemovalQueue()')
$removeQueueEnd = $c.IndexOf('        public static async Task<IEnumerable<string>> ReadPdfFileAsync', $removeQueueStart)
if ($removeQueueStart -lt 0 -or $removeQueueEnd -lt 0) { throw 'Could not locate obsolete RAG removal queue' }
$c = $c.Substring(0, $removeQueueStart) + $c.Substring($removeQueueEnd)

$old = @'
        public string GetRAGContext(string query, int maxTokens)
        {
            if (_fileList.Count == 0) return string.Empty;
            var result = _db.QueryCosineSimilarity(query, _fileList.Count * 10); // 10 results per file
            StringBuilder ragContext = new StringBuilder();
            foreach (var document in result.Documents)
                ragContext.AppendLine(document.DocumentString);
            return CommonUtils.SubstringTokens(ragContext.ToString(), maxTokens);
        }
'@
$new = @'
        public string GetRAGContext(string query, int maxTokens)
        {
            var databases = _fileDatabases.ToArray();
            if (databases.Length == 0 || maxTokens <= 0)
                return string.Empty;

            // Reserve an equal token budget for every fully indexed PDF. This prevents a
            // highly similar first paper from consuming the complete RAG context and makes
            // multi-paper comparison predictable.
            int perFileTokenBudget = Math.Max(1, maxTokens / databases.Length);
            StringBuilder ragContext = new StringBuilder();

            foreach (var entry in databases.OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase))
            {
                var result = entry.Value.QueryCosineSimilarity(query, 10);
                if (result.Documents.Count == 0)
                    continue;

                StringBuilder fileContext = new StringBuilder();
                foreach (var document in result.Documents)
                    fileContext.AppendLine(document.DocumentString);

                string boundedFileContext = CommonUtils.SubstringTokens(fileContext.ToString(), perFileTokenBudget);
                if (!string.IsNullOrWhiteSpace(boundedFileContext))
                {
                    ragContext.AppendLine(boundedFileContext);
                    ragContext.AppendLine();
                }
            }

            return CommonUtils.SubstringTokens(ragContext.ToString(), maxTokens);
        }
'@
if (-not $c.Contains($old)) { throw 'Could not locate RAG query method' }
$c = $c.Replace($old, $new)

$c = $c.Replace('try { IteratePdfFile(ref doc, ref chunks, chunkLen); }', 'try { IteratePdfFile(ref doc, ref chunks, chunkLen, Path.GetFileName(filePath)); }')
$c = $c.Replace('try { IteratePdfFile(ref unlockedDoc, ref chunks, chunkLen); }', 'try { IteratePdfFile(ref unlockedDoc, ref chunks, chunkLen, Path.GetFileName(filePath)); }')

$start = $c.IndexOf('        private static void IteratePdfFile(ref PdfDocument document, ref List<string> chunks, int chunkLen)')
$end = $c.IndexOf('        public string GetRAGContext', $start)
if ($start -lt 0 -or $end -lt 0) { throw 'Could not locate PDF iteration methods' }

$replacement = @'
        private static void IteratePdfFile(ref PdfDocument document, ref List<string> chunks, int chunkLen, string sourceName)
        {
            IterateInnerPdfFile(ref document, ref chunks, chunkLen, sourceName);

            IReadOnlyList<EmbeddedFile> embeddedFiles;
            if (document.Advanced.TryGetEmbeddedFiles(out embeddedFiles))
            {
                foreach (var embeddedFile in embeddedFiles)
                {
                    if (embeddedFile.Name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            PdfDocument embedDoc;
                            try { embedDoc = PdfDocument.Open(embeddedFile.Bytes.ToArray()); }
                            catch (PdfDocumentEncryptedException) { throw new ArgumentException(); }
                            try { IteratePdfFile(ref embedDoc, ref chunks, chunkLen, embeddedFile.Name); }
                            finally { embedDoc.Dispose(); }
                        }
                        catch (ArgumentException)
                        {
                            PasswordPrompt passwordDialog = new PasswordPrompt();
                            if (passwordDialog.ShowDialog() == DialogResult.OK)
                            {
                                PdfDocument unlockedDoc = PdfDocument.Open(embeddedFile.Bytes.ToArray(), new ParsingOptions { Password = passwordDialog.Password });
                                try { IteratePdfFile(ref unlockedDoc, ref chunks, chunkLen, embeddedFile.Name); }
                                finally { unlockedDoc.Dispose(); }
                            }
                            else
                            {
                                throw new InvalidDataException(_cultureHelper.GetLocalizedString("[ReadPdfFileAsync] InvalidDataException #1"));
                            }
                        }
                    }
                }
            }
        }

        private static void IterateInnerPdfFile(ref PdfDocument doc, ref List<string> chunks, int chunkLen, string sourceName)
        {
            foreach (var page in doc.GetPages())
            {
                var pageText = new StringBuilder();
                var blocks = DocstrumBoundingBoxes.Instance.GetBlocks(page.GetWords());
                foreach (var block in blocks)
                {
                    string text = NormalizePdfText(block.Text);
                    if (string.IsNullOrWhiteSpace(text))
                        continue;
                    if (IsLowValueStandaloneBlock(text))
                        continue;
                    if (pageText.Length > 0)
                        pageText.AppendLine();
                    pageText.Append(text);
                }

                string normalizedPage = NormalizePdfText(pageText.ToString());
                if (normalizedPage.Length < 80)
                    continue;

                foreach (string chunk in SplitTextWithOverlap(normalizedPage, chunkLen, CHUNK_OVERLAP))
                {
                    string metadata = $"[Source: {sourceName}; Page: {page.Number}]";
                    chunks.Add(metadata + Environment.NewLine + chunk);
                }
            }
        }

        private static string NormalizePdfText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;
            text = text.Replace("\u00AD", string.Empty);
            text = Regex.Replace(text, @"\s+", " ");
            return text.Trim();
        }

        private static bool IsLowValueStandaloneBlock(string text)
        {
            if (text.Length > 80)
                return false;
            string value = text.Trim().Trim(':').ToLowerInvariant();
            string[] lowValue =
            {
                "abstract", "references", "bibliography", "supporting information",
                "supplementary information", "acknowledgements", "acknowledgments",
                "keywords", "contents"
            };
            return lowValue.Contains(value);
        }

        private static IEnumerable<string> SplitTextWithOverlap(string text, int chunkLen, int overlapLen)
        {
            if (string.IsNullOrWhiteSpace(text))
                yield break;
            chunkLen = Math.Max(256, chunkLen);
            overlapLen = Math.Max(0, Math.Min(overlapLen, chunkLen / 2));
            int start = 0;
            while (start < text.Length)
            {
                int remaining = text.Length - start;
                int length = Math.Min(chunkLen, remaining);
                int end = start + length;
                if (end < text.Length)
                {
                    int searchStart = start + (int)(length * 0.65);
                    int sentenceEnd = -1;
                    for (int i = end - 1; i >= searchStart; i--)
                    {
                        char ch = text[i];
                        if (ch == '.' || ch == '!' || ch == '?' || ch == ';')
                        {
                            sentenceEnd = i + 1;
                            break;
                        }
                    }
                    if (sentenceEnd > start)
                        end = sentenceEnd;
                }
                string chunk = text.Substring(start, end - start).Trim();
                if (chunk.Length >= 80)
                    yield return chunk;
                if (end >= text.Length)
                    break;
                int nextStart = Math.Max(start + 1, end - overlapLen);
                while (nextStart < end && nextStart < text.Length && !char.IsWhiteSpace(text[nextStart]))
                    nextStart++;
                start = nextStart;
            }
        }

'@
$c = $c.Substring(0, $start) + $replacement + $c.Substring($end)
Set-Content $path $c -Encoding UTF8

Write-Host 'Patching ThisAddIn.cs...'
$path = 'ThisAddIn.cs'
$c = Get-Content $path -Raw
$old = @'
        private void RemoveClosedWindowTaskPanes()
        {
            for (int i = this.CustomTaskPanes.Count - 1; i >= 0; i--)
                if (this.CustomTaskPanes[i].Window == null)
                    this.CustomTaskPanes.RemoveAt(i);
        }
'@
$new = @'
        private void RemoveClosedWindowTaskPanes()
        {
            for (int i = this.CustomTaskPanes.Count - 1; i >= 0; i--)
            {
                bool removePane = false;
                try { removePane = this.CustomTaskPanes[i].Window == null; }
                catch (System.Runtime.InteropServices.COMException) { removePane = true; }
                catch (ObjectDisposedException) { removePane = true; }
                if (!removePane) continue;
                try { this.CustomTaskPanes.RemoveAt(i); }
                catch (System.Runtime.InteropServices.COMException) { }
                catch (ObjectDisposedException) { }
                catch (ArgumentOutOfRangeException) { }
            }
        }
'@
if (-not $c.Contains($old)) { throw 'Could not locate task pane cleanup method' }
$c = $c.Replace($old, $new)
Set-Content $path $c -Encoding UTF8

Write-Host 'Restoring NuGet packages...'
nuget restore TextForge.sln -NonInteractive
if ($LASTEXITCODE -ne 0) { throw "NuGet restore failed: $LASTEXITCODE" }

Write-Host 'Building TextCraft Release...'
msbuild TextCraft.csproj /m /t:Rebuild /p:Configuration=Release /p:Platform=AnyCPU
if ($LASTEXITCODE -ne 0) { throw "MSBuild failed: $LASTEXITCODE" }

Write-Host 'Trying MSI build...'
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$vs = & $vswhere -latest -products * -property installationPath
$devenv = Join-Path $vs 'Common7\IDE\devenv.com'
& $devenv TextForge.sln /Build 'Release|Any CPU' /Project OfficeAddInSetup
$msiExit = $LASTEXITCODE
if ($msiExit -ne 0) { Write-Warning "MSI build failed with exit code $msiExit; binary build will still be uploaded." }

New-Item -ItemType Directory -Force artifact | Out-Null
if (Test-Path 'bin\Release') { Copy-Item 'bin\Release\*' artifact -Recurse -Force }

# VSTO/ClickOnce manifest includes this System.Text.Json linker descriptor as an application file.
# MSBuild leaves it in the restored package tree instead of bin\Release, so copy it explicitly.
$ilLinkDescriptor = Get-ChildItem 'packages' -Recurse -File -Filter 'ILLink.Descriptors.LibraryBuild.xml' -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $ilLinkDescriptor) { throw 'ILLink.Descriptors.LibraryBuild.xml was referenced by the VSTO manifest but not found in restored packages.' }
New-Item -ItemType Directory -Force 'artifact\ILLink' | Out-Null
Copy-Item $ilLinkDescriptor.FullName 'artifact\ILLink\ILLink.Descriptors.LibraryBuild.xml' -Force
Write-Host "Included VSTO application file: $($ilLinkDescriptor.FullName)"

Get-ChildItem 'OfficeAddInSetup' -Recurse -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Extension -in '.msi','.exe' } |
    Copy-Item -Destination artifact -Force
Copy-Item RAGControl.cs artifact\RAGControl.patched.cs
Copy-Item ThisAddIn.cs artifact\ThisAddIn.patched.cs

Write-Host 'Build preparation complete.'
# CI packaging revision: trust-certificate bundle is added by the workflow.
