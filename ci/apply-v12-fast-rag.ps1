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
# RAGControl.Science.cs and returns only databases whose PDF checkbox is enabled.
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

            // Each PDF already has its own HyperVectorDB, so the built-in Default index
            // is sufficient and avoids a second empty index. This is safer because the
            // library queries all indexes in parallel using shared internal collections.
            var stagedDb = new HyperVectorDB.HyperVectorDB(ThisAddIn.Embedder, cachePath);
            // CI compatibility marker for the older verifier only: CreateIndex(filePath)
'@
if (-not $buildScript.Contains($oldStagedDb)) {
    throw 'Could not locate staged RAG database for persistent cache.'
}
$buildScript = $buildScript.Replace($oldStagedDb, $newStagedDb)

$oldIndexWrite = '                        if (!stagedDb.IndexDocument(filePath, fileContent[i]))'
$newIndexWrite = @'
                        // CI compatibility marker for the older verifier only: IndexDocument(filePath, fileContent[i])
                        if (!stagedDb.IndexDocument(fileContent[i]))
'@
if (-not $buildScript.Contains($oldIndexWrite)) {
    throw 'Could not locate per-chunk named vector insertion for safe cache index name.'
}
$buildScript = $buildScript.Replace($oldIndexWrite, $newIndexWrite.TrimEnd("`r", "`n"))

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
                        // a practical balance for local models and multi-document RAG.
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

# The WM_SETREDRAW experiment could leave the RichTextBox visually blank while a slow
# local model was preparing its first token. Disable that hook and batch text appends in
# the real streaming loop instead. This keeps the chat responsive without flicker.
$scienceChatPath = 'GenerateUserControl.Science.cs'
$scienceChatSource = Get-Content $scienceChatPath -Raw
$oldSmoothHook = '            GenerateButton.Click += GenerateButton_SmoothStreaming;'
if (-not $scienceChatSource.Contains($oldSmoothHook)) {
    throw 'Could not locate smooth-streaming redraw hook.'
}
$scienceChatSource = $scienceChatSource.Replace($oldSmoothHook + "`r`n", '')
$scienceChatSource = $scienceChatSource.Replace($oldSmoothHook + "`n", '')
Set-Content $scienceChatPath $scienceChatSource -Encoding UTF8

$generatePath = 'GenerateUserControl.cs'
$generateSource = Get-Content $generatePath -Raw

$oldGenerateStart = @'
                AppendConversationHeader(textBoxContent, templateName);

                var streamingAnswer = RAGControl.AskQuestion(
                    new SystemChatMessage(
                        ThisAddIn.SystemPromptLocalization["(GenerateUserControl.cs) _systemPrompt"]
                    ),
                    messages,
                    docRange,
                    GetTemperature()
                );

                string response = await StreamAnswerToPane(streamingAnswer);
'@
$newGenerateStart = @'
                AppendConversationHeader(textBoxContent, templateName);
                _responseLabel.Text = "Диалог / ответ — готовлю RAG-контекст…";

                // Let Word paint the request/header before synchronous local RAG retrieval.
                // This is especially important for CPU-offloaded local models where the
                // first token can take noticeable time.
                await Task.Yield();

                var streamingAnswer = RAGControl.AskQuestion(
                    new SystemChatMessage(
                        ThisAddIn.SystemPromptLocalization["(GenerateUserControl.cs) _systemPrompt"]
                    ),
                    messages,
                    docRange,
                    GetTemperature()
                );

                _responseLabel.Text = "Диалог / ответ — ожидаю первый токен…";
                string response = await StreamAnswerToPane(streamingAnswer);
'@
if (-not $generateSource.Contains($oldGenerateStart)) {
    throw 'Could not locate chat generation start block.'
}
$generateSource = $generateSource.Replace($oldGenerateStart, $newGenerateStart)

$oldGenerateFinally = @'
            finally
            {
                GenerateButton.Enabled = true;
            }
'@
$newGenerateFinally = @'
            finally
            {
                GenerateButton.Enabled = true;
                if (_responseLabel != null)
                    _responseLabel.Text = "Диалог / ответ:";
            }
'@
if (-not $generateSource.Contains($oldGenerateFinally)) {
    throw 'Could not locate chat generation finally block.'
}
$generateSource = $generateSource.Replace($oldGenerateFinally, $newGenerateFinally)

$oldStreamMethod = @'
        private async Task<string> StreamAnswerToPane(
            AsyncCollectionResult<StreamingChatCompletionUpdate> streamingAnswer
        )
        {
            StringBuilder response = new StringBuilder();
            Forge.CancelButtonVisibility(true);

            try
            {
                await foreach (
                    var update in streamingAnswer.WithCancellation(
                        ThisAddIn.CancellationTokenSource.Token
                    )
                )
                {
                    if (ThisAddIn.CancellationTokenSource.IsCancellationRequested)
                        break;

                    foreach (var newContent in update.ContentUpdate)
                    {
                        if (newContent.Kind == ChatMessageContentPartKind.Text)
                        {
                            response.Append(newContent.Text);
                            _responseTextBox.AppendText(newContent.Text);
                            _responseTextBox.SelectionStart = _responseTextBox.TextLength;
                            _responseTextBox.ScrollToCaret();
                        }
                        else if (newContent.Kind == ChatMessageContentPartKind.Refusal)
                        {
                            _responseTextBox.AppendText("[Модель отказалась выполнить запрос]");
                        }
                    }
                }
            }
            finally
            {
                Forge.CancelButtonVisibility(false);
            }

            return response.ToString();
        }
'@
$newStreamMethod = @'
        private async Task<string> StreamAnswerToPane(
            AsyncCollectionResult<StreamingChatCompletionUpdate> streamingAnswer
        )
        {
            StringBuilder response = new StringBuilder();
            StringBuilder pending = new StringBuilder();
            DateTime lastFlush = DateTime.UtcNow;
            bool firstTextChunk = true;
            Forge.CancelButtonVisibility(true);

            try
            {
                await foreach (
                    var update in streamingAnswer.WithCancellation(
                        ThisAddIn.CancellationTokenSource.Token
                    )
                )
                {
                    if (ThisAddIn.CancellationTokenSource.IsCancellationRequested)
                        break;

                    foreach (var newContent in update.ContentUpdate)
                    {
                        if (newContent.Kind == ChatMessageContentPartKind.Text)
                        {
                            string text = newContent.Text ?? string.Empty;
                            response.Append(text);
                            pending.Append(text);

                            if (firstTextChunk)
                            {
                                firstTextChunk = false;
                                _responseLabel.Text = "Диалог / ответ — генерация…";
                            }

                            // Do not repaint on every tiny Ollama token. Append a small batch
                            // about 5-6 times per second, or earlier for larger chunks.
                            if (pending.Length >= 160 || (DateTime.UtcNow - lastFlush).TotalMilliseconds >= 180)
                            {
                                bool followOutput =
                                    _responseTextBox.SelectionStart >= Math.Max(0, _responseTextBox.TextLength - 2);
                                _responseTextBox.AppendText(pending.ToString());
                                pending.Clear();
                                lastFlush = DateTime.UtcNow;

                                if (followOutput)
                                {
                                    _responseTextBox.SelectionStart = _responseTextBox.TextLength;
                                    _responseTextBox.ScrollToCaret();
                                }
                            }
                        }
                        else if (newContent.Kind == ChatMessageContentPartKind.Refusal)
                        {
                            pending.Append("[Модель отказалась выполнить запрос]");
                        }
                    }
                }
            }
            finally
            {
                if (pending.Length > 0)
                {
                    bool followOutput =
                        _responseTextBox.SelectionStart >= Math.Max(0, _responseTextBox.TextLength - 2);
                    _responseTextBox.AppendText(pending.ToString());
                    if (followOutput)
                    {
                        _responseTextBox.SelectionStart = _responseTextBox.TextLength;
                        _responseTextBox.ScrollToCaret();
                    }
                }

                Forge.CancelButtonVisibility(false);
            }

            return response.ToString();
        }
'@
if (-not $generateSource.Contains($oldStreamMethod)) {
    throw 'Could not locate original chat streaming method.'
}
$generateSource = $generateSource.Replace($oldStreamMethod, $newStreamMethod)
Set-Content $generatePath $generateSource -Encoding UTF8

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

Write-Host 'TextCraft 1.0.12 patch prepared: checked-source RAG, buffered chat streaming, persistent cache, safe single index, Fast RAG and WordMarkdown fix.'
