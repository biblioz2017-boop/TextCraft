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

Write-Host 'Applying automatic RAG topic discovery...'

# -----------------------------------------------------------------------------
# RAGControl: build a broad semantic seed directly from the checked PDFs.
# This does not use model knowledge. It reads a small sample from the first pages
# of the files that the user actually checked in the Literature pane.
# -----------------------------------------------------------------------------
$ragPath = 'RAGControl.cs'
$rag = Read-Utf8Text $ragPath
$rnl = Get-NewLine $rag

if (-not $rag.Contains('public async Task<string> BuildCheckedPdfTopicSeedAsync()')) {
    $anchor = '        // UTILS'
    if (-not $rag.Contains($anchor)) {
        throw 'Could not locate RAGControl utility anchor for topic discovery.'
    }

    $methods = @'
        public async Task<string> BuildCheckedPdfTopicSeedAsync()
        {
            var checkedFiles = new List<string>();

            Action snapshot = () =>
            {
                if (FileListBox == null || FileListBox.IsDisposed)
                    return;

                foreach (object rawItem in FileListBox.CheckedItems)
                {
                    if (rawItem is KeyValuePair<string, string> item &&
                        !string.IsNullOrWhiteSpace(item.Value))
                    {
                        checkedFiles.Add(item.Value);
                    }
                }
            };

            if (FileListBox != null && !FileListBox.IsDisposed && FileListBox.InvokeRequired)
                FileListBox.Invoke(snapshot);
            else
                snapshot();

            if (checkedFiles.Count == 0)
                return string.Empty;

            return await Task.Run(() => BuildCheckedPdfTopicSeed(checkedFiles));
        }

        private static string BuildCheckedPdfTopicSeed(List<string> checkedFiles)
        {
            var seed = new StringBuilder();
            const int maxFiles = 8;
            const int maxPagesPerFile = 2;
            const int maxWordsPerPage = 180;
            const int maxSeedCharacters = 6000;

            int fileCount = 0;
            foreach (string filePath in checkedFiles)
            {
                if (fileCount >= maxFiles || seed.Length >= maxSeedCharacters)
                    break;
                if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                    continue;

                try
                {
                    using (PdfDocument document = PdfDocument.Open(filePath))
                    {
                        int pageLimit = Math.Min(maxPagesPerFile, document.NumberOfPages);
                        for (int pageNumber = 1; pageNumber <= pageLimit; pageNumber++)
                        {
                            Page page = document.GetPage(pageNumber);
                            int wordCount = 0;
                            foreach (Word word in page.GetWords())
                            {
                                if (word == null || string.IsNullOrWhiteSpace(word.Text))
                                    continue;

                                seed.Append(word.Text).Append(' ');
                                wordCount++;
                                if (wordCount >= maxWordsPerPage || seed.Length >= maxSeedCharacters)
                                    break;
                            }

                            seed.AppendLine();
                            if (seed.Length >= maxSeedCharacters)
                                break;
                        }
                    }
                    fileCount++;
                }
                catch
                {
                    // One unreadable PDF must not block topic discovery from the others.
                }
            }

            string result = seed.ToString().Trim();
            if (result.Length > maxSeedCharacters)
                result = result.Substring(0, maxSeedCharacters);
            return result;
        }

'@
    $rag = $rag.Replace($anchor, $methods.Replace("`n", $rnl) + $anchor)
}

Write-Utf8Text $ragPath $rag

# -----------------------------------------------------------------------------
# Science pane: when the prompt is only "make a report from attached files",
# use the checked-PDF seed as the retrieval query. Specific topical prompts keep
# their current precise semantic retrieval behavior.
# -----------------------------------------------------------------------------
$sciencePath = 'GenerateUserControl.Science.cs'
$science = Read-Utf8Text $sciencePath
$snl = Get-NewLine $science

if (-not $science.Contains('private static bool IsGenericAttachedSourceRequest(string userQuery)')) {
    $anchor = '        private string BuildRagRetrievalQuery(string userQuery)'
    if (-not $science.Contains($anchor)) {
        throw 'Could not locate BuildRagRetrievalQuery for generic-source detection.'
    }

    $helper = @'
        private static bool IsGenericAttachedSourceRequest(string userQuery)
        {
            string normalized = Regex.Replace(
                (userQuery ?? string.Empty).ToLowerInvariant(),
                @"[^\p{L}\p{Nd}]+",
                " "
            ).Trim();

            if (normalized.Length == 0 || normalized.Length > 220)
                return false;

            string[] sourcePrefixes =
            {
                "\u0444\u0430\u0439\u043b",
                "pdf",
                "\u043f\u0440\u0438\u043a\u0440\u0435\u043f",
                "\u043e\u0442\u043c\u0435\u0447",
                "\u0438\u0441\u0442\u043e\u0447\u043d",
                "\u043c\u0430\u0442\u0435\u0440\u0438\u0430\u043b",
                "\u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442"
            };

            string[] genericPrefixes =
            {
                "\u0441\u0434\u0435\u043b",
                "\u043f\u043e\u0434\u0433\u043e\u0442\u043e\u0432",
                "\u043d\u0430\u043f\u0438\u0448",
                "\u0441\u043e\u0441\u0442\u0430\u0432",
                "\u0440\u0435\u0444\u0435\u0440\u0430\u0442",
                "\u043e\u0431\u0437\u043e\u0440",
                "\u043a\u043e\u043d\u0441\u043f\u0435\u043a\u0442",
                "\u043f\u0440\u0438\u043a\u0440\u0435\u043f",
                "\u043e\u0442\u043c\u0435\u0447",
                "\u0444\u0430\u0439\u043b",
                "pdf",
                "\u0438\u0441\u0442\u043e\u0447\u043d",
                "\u043c\u0430\u0442\u0435\u0440\u0438\u0430\u043b",
                "\u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442",
                "\u043e\u0441\u043d\u043e\u0432",
                "\u0438\u0441\u043f\u043e\u043b\u044c\u0437",
                "\u0434\u0430\u043d\u043d"
            };

            string[] words = normalized.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            bool mentionsSources = words.Any(
                word => sourcePrefixes.Any(prefix => word.StartsWith(prefix, StringComparison.Ordinal))
            );
            if (!mentionsSources)
                return false;

            int meaningfulWords = 0;
            foreach (string word in words)
            {
                if (word.Length < 4)
                    continue;
                if (genericPrefixes.Any(prefix => word.StartsWith(prefix, StringComparison.Ordinal)))
                    continue;

                meaningfulWords++;
                if (meaningfulWords > 2)
                    return false;
            }

            return true;
        }

'@
    $science = $science.Replace($anchor, $helper.Replace("`n", $snl) + $anchor)
}

$oldEvidenceCall = '                    forcedEvidence = await Task.Run(() => rag.GetRAGEvidence(retrievalQuery, 4));'
if ($science.Contains($oldEvidenceCall)) {
    $newEvidenceCall = @'
                    string evidenceQuery = retrievalQuery;
                    if (IsGenericAttachedSourceRequest(userQuery))
                    {
                        if (_responseLabel != null)
                            _responseLabel.Text = "\u0414\u0438\u0430\u043b\u043e\u0433 / \u043e\u0442\u0432\u0435\u0442 \u2014 \u043e\u043f\u0440\u0435\u0434\u0435\u043b\u044f\u044e \u0442\u0435\u043c\u0443 \u043f\u043e \u043e\u0442\u043c\u0435\u0447\u0435\u043d\u043d\u044b\u043c PDF\u2026";

                        string sourceTopicSeed = await rag.BuildCheckedPdfTopicSeedAsync();
                        if (!string.IsNullOrWhiteSpace(sourceTopicSeed))
                            evidenceQuery = sourceTopicSeed;
                    }

                    forcedEvidence = await Task.Run(() => rag.GetRAGEvidence(evidenceQuery, 6));
'@
    $science = $science.Replace($oldEvidenceCall, $newEvidenceCall.TrimEnd("`r", "`n").Replace("`n", $snl))
} elseif (-not $science.Contains('string sourceTopicSeed = await rag.BuildCheckedPdfTopicSeedAsync();')) {
    throw 'Could not locate strict-RAG evidence retrieval call for auto-topic fallback.'
}

Write-Utf8Text $sciencePath $science

if (-not (Read-Utf8Text $ragPath).Contains('BuildCheckedPdfTopicSeedAsync')) {
    throw 'Checked-PDF topic discovery method is missing.'
}
if (-not (Read-Utf8Text $sciencePath).Contains('IsGenericAttachedSourceRequest')) {
    throw 'Generic attached-source request detector is missing.'
}
if (-not (Read-Utf8Text $sciencePath).Contains('sourceTopicSeed = await rag.BuildCheckedPdfTopicSeedAsync')) {
    throw 'Automatic topic fallback is not wired into strict RAG.'
}

Write-Host 'Automatic RAG topic discovery applied successfully.'
