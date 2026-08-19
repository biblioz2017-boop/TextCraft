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

Write-Host 'TextCraft 1.0.12 Fast RAG patch prepared: Ollama context capped at 8192, top-4 chunks per PDF.'
