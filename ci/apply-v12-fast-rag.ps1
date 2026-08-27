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

# Fix WordMarkdown absolute/relative index mixing. Match.Index is already the
# correct position inside partialMarkdownText; feeding an absolute Word position
# back into String.IndexOf(startIndex) can throw ArgumentOutOfRangeException.
$markdownPath = 'WordMarkdown.cs'
$markdownSource = Get-Content $markdownPath -Raw

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

Set-Content $markdownPath $markdownSource -Encoding UTF8

Write-Host 'TextCraft 1.0.12 patch prepared: Fast RAG plus WordMarkdown absolute-index fix.'
