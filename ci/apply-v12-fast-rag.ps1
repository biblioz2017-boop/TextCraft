$ErrorActionPreference = 'Stop'

# Keep all verified 1.0.11 vector-indexing fixes.
./ci/apply-v11-fix.ps1

$buildPath = 'ci/build-patched.ps1'
$buildScript = Get-Content $buildPath -Raw

$oldQuery = '                var result = entry.Value.QueryCosineSimilarity(query, 10);'
$newQuery = '                var result = entry.Value.QueryCosineSimilarity(query, 4);'
if (-not $buildScript.Contains($oldQuery)) {
    throw 'Could not locate per-PDF RAG retrieval count for Fast RAG patch.'
}
$buildScript = $buildScript.Replace($oldQuery, $newQuery)

# Respect the source subset selected in the Literature pane. The helper lives in
# RAGControl.Science.cs and returns all databases when source filtering is disabled.
$oldDatabases = '            var databases = _fileDatabases.ToArray();'
$newDatabases = '            var databases = GetActiveRagDatabases();'
if (-not $buildScript.Contains($oldDatabases)) {
    throw 'Could not locate active RAG database selection.'
}
$buildScript = $buildScript.Replace($oldDatabases, $newDatabases)

# Reuse a persistent vector index when the exact same PDF, chunk settings and
# embedding model were indexed previously.
$oldIndexStart = @'
            long indexVersion = Interlocked.Increment(ref _nextIndexVersion);
            _activeIndexVersions[filePath] = indexVersion;

            List<string> fileContent;
'@
$newIndexStart = @'
            long indexVersion = Interlocked.Increment(ref _nextIndexVersion);
            _activeIndexVersions[filePath] = indexVersion;

            HyperVectorDB.HyperVectorDB cachedDb;
            if (TryLoadCachedDatabase(filePath, out cachedDb))
            {
                _fileDatabases[filePath] = cachedDb;
                MarkFileStatus(filePath, $"[CACHE] {Path.GetFileName(filePath)}");
                return;
            }

            List<string> fileContent;
'@
if (-not $buildScript.Contains($oldIndexStart)) {
    throw 'Could not locate RAG indexing start for persistent cache.'
}
$buildScript = $buildScript.Replace($oldIndexStart, $newIndexStart)

$oldStagedDb = @'
            var stagedDb = new HyperVectorDB.HyperVectorDB(ThisAddIn.Embedder, Path.GetTempPath());
            if (!stagedDb.CreateIndex(filePath))
                throw new InvalidOperationException($"Could not create vector index for {filePath}");
'@
$newStagedDb = @'
            string cachePath = GetPersistentDatabasePath(filePath);
            if (Directory.Exists(cachePath))
            {
                try { Directory.Delete(cachePath, true); }
                catch { }
            }

            var stagedDb = new HyperVectorDB.HyperVectorDB(ThisAddIn.Embedder, cachePath);
            if (!stagedDb.CreateIndex(filePath))
                throw new InvalidOperationException($"Could not create vector index for {filePath}");
'@
if (-not $buildScript.Contains($oldStagedDb)) {
    throw 'Could not locate staged RAG database for persistent cache.'
}
$buildScript = $buildScript.Replace($oldStagedDb, $newStagedDb)

$oldPublishDb = '                _fileDatabases[filePath] = stagedDb;'
$newPublishDb = "                stagedDb.Save();`r`n                _fileDatabases[filePath] = stagedDb;"
if (-not $buildScript.Contains($oldPublishDb)) {
    throw 'Could not locate staged RAG publish step.'
}
$buildScript = $buildScript.Replace($oldPublishDb, $newPublishDb)

Set-Content $buildPath $buildScript -Encoding UTF8

$modelPath = 'ModelProperties.cs'
$modelSource = Get-Content $modelPath -Raw
$oldContext = @'
                        return contextWindow;
'@
$newContext = @'
                        // Fast local-RAG cap: large advertised Ollama context windows can
                        // cause TextCraft to allocate excessive prompt/RAG budgets. 8K is
                        // a practical balance for 4B local models and multi-document RAG.
                        return Math.Min(contextWindow, 8192);
'@
if (-not $modelSource.Contains($oldContext)) {
    throw 'Could not locate Ollama context-window return for Fast RAG patch.'
}
$modelSource = $modelSource.Replace($oldContext, $newContext)
Set-Content $modelPath $modelSource -Encoding UTF8

# Quick rewrite actions should use Word's native revision tracking so the user can
# accept or reject AI edits with the standard Review tab.
$forgeDesignerPath = 'Forge.Designer.cs'
$forgeDesignerSource = Get-Content $forgeDesignerPath -Raw
$oldQuickEdit = '                await AnalyzeText(QuickTextSystemPrompt, instruction, temperature);'
$newQuickEdit = '                await AnalyzeTextWithTrackChanges(QuickTextSystemPrompt, instruction, temperature);'
if (-not $forgeDesignerSource.Contains($oldQuickEdit)) {
    throw 'Could not locate quick text action for Track Changes wrapper.'
}
$forgeDesignerSource = $forgeDesignerSource.Replace($oldQuickEdit, $newQuickEdit)
Set-Content $forgeDesignerPath $forgeDesignerSource -Encoding UTF8

# Compile the scientific workflow partial classes without changing the upstream
# project structure more than necessary.
$projectPath = 'TextCraft.csproj'
$projectSource = Get-Content $projectPath -Raw
$compileAnchor = '    <Compile Include="ModelProperties.cs" />'
$compileInsert = @'
    <Compile Include="Forge.Science.cs" />
    <Compile Include="GenerateUserControl.Science.cs" />
    <Compile Include="RAGControl.Science.cs" />
    <Compile Include="ModelProperties.cs" />
'@
if (-not $projectSource.Contains($compileAnchor)) {
    throw 'Could not locate project compile anchor for scientific workflow files.'
}
$projectSource = $projectSource.Replace($compileAnchor, $compileInsert)
Set-Content $projectPath $projectSource -Encoding UTF8

# Use Russian task-pane captions matching the simplified ribbon.
$thisAddInPath = 'ThisAddIn.cs'
$thisAddInSource = Get-Content $thisAddInPath -Raw
$thisAddInSource = $thisAddInSource.Replace(
    'Globals.ThisAddIn.CustomTaskPanes.Add(new GenerateUserControl(), Forge.CultureHelper.GetLocalizedString("this.GenerateButton.Label"), doc.ActiveWindow)',
    'Globals.ThisAddIn.CustomTaskPanes.Add(new GenerateUserControl(), "Спросить", doc.ActiveWindow)'
)
$thisAddInSource = $thisAddInSource.Replace(
    'Globals.ThisAddIn.CustomTaskPanes.Add(ragControl, Forge.CultureHelper.GetLocalizedString("this.RAGControlButton.Label"), doc.ActiveWindow)',
    'Globals.ThisAddIn.CustomTaskPanes.Add(ragControl, "Литература", doc.ActiveWindow)'
)
Set-Content $thisAddInPath $thisAddInSource -Encoding UTF8

# Fix WordMarkdown absolute/relative index mixing. Match.Index is already the
# correct position inside partialMarkdownText; feeding an absolute Word position
# back into String.IndexOf(startIndex) can throw ArgumentOutOfRangeException.
$markdownPath = 'WordMarkdown.cs'
$markdownSource = Get-Content $markdownPath -Raw

# IMPORTANT: patch GetCodeBlockPoints first. Its opening block is intentionally
# similar to ApplyMarkdownFormatting, so a broad Replace on the latter first would
# also rewrite this method and make the more specific replacement impossible.
$oldCodeBlockPoints = @'
            int searchIndex = 0;
            int offset = 0;
            foreach (Match match in matches)
            {
                string textToFormat = match.Value;
                string insideContent = match.Groups[1].Value;
                searchIndex = commentRange.Start + partialMarkdownText.IndexOf(match.Value, searchIndex);
                int length = textToFormat.Length;

                points.Add(new CodeBlockPoint(searchIndex - offset, searchIndex - offset + length - 1, insideContent.Length));
'@
$newCodeBlockPoints = @'
            int offset = 0;
            foreach (Match match in matches)
            {
                string textToFormat = match.Value;
                string insideContent = match.Groups[1].Value;
                int searchIndex = commentRange.Start + match.Index;
                int length = textToFormat.Length;

                points.Add(new CodeBlockPoint(searchIndex - offset, searchIndex - offset + length - 1, insideContent.Length));
'@
if (-not $markdownSource.Contains($oldCodeBlockPoints)) {
    throw 'Could not locate GetCodeBlockPoints searchIndex block.'
}
$markdownSource = $markdownSource.Replace($oldCodeBlockPoints, $newCodeBlockPoints)

$oldApplyHead = @'
            int searchIndex = 0;
            int offset = 0;
            foreach (Match match in matches)
            {
                string textToFormat = match.Value;
                string insideContent = match.Groups[1].Value;
                searchIndex = commentRange.Start + partialMarkdownText.IndexOf(match.Value, searchIndex);
                int length = textToFormat.Length;
'@
$newApplyHead = @'
            int offset = 0;
            foreach (Match match in matches)
            {
                string textToFormat = match.Value;
                string insideContent = match.Groups[1].Value;
                int searchIndex = commentRange.Start + match.Index;
                int length = textToFormat.Length;
'@
if (-not $markdownSource.Contains($oldApplyHead)) {
    throw 'Could not locate ApplyMarkdownFormatting searchIndex block.'
}
$markdownSource = $markdownSource.Replace($oldApplyHead, $newApplyHead)

$oldApplyTail = @'
                }
                searchIndex += length;
            }
        }

        // Add method to handle LaTeX equations:
'@
$newApplyTail = @'
                }
            }
        }

        // Add method to handle LaTeX equations:
'@
if (-not $markdownSource.Contains($oldApplyTail)) {
    throw 'Could not locate ApplyMarkdownFormatting searchIndex increment.'
}
$markdownSource = $markdownSource.Replace($oldApplyTail, $newApplyTail)

Set-Content $markdownPath $markdownSource -Encoding UTF8

Write-Host 'TextCraft 1.0.12 patch prepared: scientific RAG workflow, persistent cache, Fast RAG and WordMarkdown fix.'
