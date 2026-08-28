$ErrorActionPreference = 'Stop'

# This patch runs immediately before C# compilation (via Directory.Build.targets).
# It is intentionally idempotent because MSBuild can invoke BeforeCompile more than once.

Write-Host 'Applying RAG chat semantic-query patch...'
$ragPath = 'RAGControl.cs'
$rag = Get-Content $ragPath -Raw

if (-not $rag.Contains('string semanticRagQuery = lastUserPrompt.Content[0].Text;')) {
    $old = @'
            ChatMessage lastUserPrompt = messages.Last();

            var constraints = RAGControl.OptimizeConstraint(
'@
    $new = @'
            ChatMessage lastUserPrompt = messages.Last();

            // The visible chat prompt can contain instructions and previous-turn context.
            // For vector retrieval use only the explicit scientific topic written by the
            // RAG-aware chat UI. This prevents follow-ups such as "дополни из RAG" from
            // searching for generic command words instead of the actual subject.
            string semanticRagQuery = lastUserPrompt.Content[0].Text;
            const string ragTopicMarker = "Тема для семантического поиска в RAG:";
            int ragTopicStart = semanticRagQuery.IndexOf(ragTopicMarker, StringComparison.OrdinalIgnoreCase);
            if (ragTopicStart >= 0)
            {
                ragTopicStart += ragTopicMarker.Length;
                int ragTopicEnd = semanticRagQuery.IndexOf('\n', ragTopicStart);
                if (ragTopicEnd < 0)
                    ragTopicEnd = semanticRagQuery.Length;

                string extractedTopic = semanticRagQuery.Substring(ragTopicStart, ragTopicEnd - ragTopicStart).Trim();
                if (!string.IsNullOrWhiteSpace(extractedTopic))
                    semanticRagQuery = extractedTopic;
            }

            var constraints = RAGControl.OptimizeConstraint(
'@

    if (-not $rag.Contains($old)) {
        throw 'Could not locate ProcessInformation semantic-query insertion point.'
    }
    $rag = $rag.Replace($old, $new)
}

$rag = $rag.Replace(
    'document = RAGControl.GetWordDocumentAsRAG(lastUserPrompt.Content[0].Text, context);',
    'document = RAGControl.GetWordDocumentAsRAG(semanticRagQuery, context);'
)

$rag = $rag.Replace(
    'ThisAddIn.AllTaskPanes[doc].Item3.GetRAGContext(lastUserPrompt.Content[0].Text, (int)(ThisAddIn.ContextLength * constraints["rag_context"]))',
    'ThisAddIn.AllTaskPanes[doc].Item3.GetRAGContext(semanticRagQuery, (int)(ThisAddIn.ContextLength * constraints["rag_context"]))'
)

Set-Content $ragPath $rag -Encoding UTF8

Write-Host 'Applying explicit default-model UI patch...'
$designerPath = 'Forge.Designer.cs'
$designer = Get-Content $designerPath -Raw

# The simplified UI originally hid the upstream default-model checkbox and did not
# place it in the Model group. Restore it as an explicit user action.
if (-not $designer.Contains('this.SettingsGroup.Items.Add(this.DefaultCheckBox);')) {
    $anchor = '            this.SettingsGroup.Items.Add(this.ModelListDropDown);'
    if (-not $designer.Contains($anchor)) {
        throw 'Could not locate SettingsGroup model dropdown.'
    }
    $designer = $designer.Replace(
        $anchor,
        $anchor + "`r`n            this.SettingsGroup.Items.Add(this.DefaultCheckBox);"
    )
}

$designer = $designer.Replace('            this.DefaultCheckBox.Visible = false;', '            this.DefaultCheckBox.Visible = true;')
$designer = $designer.Replace('            this.DefaultCheckBox.Label = "По умолчанию";', '            this.DefaultCheckBox.Label = "По умолчанию";')
Set-Content $designerPath $designer -Encoding UTF8

$forgePath = 'Forge.cs'
$forge = Get-Content $forgePath -Raw

# Selecting a model now changes only the current session. It becomes the startup model
# only after the user explicitly checks "По умолчанию".
$oldSelectionSave = @'
                ThisAddIn.Model = selectedModel;

                Properties.Settings.Default.DefaultModel = selectedModel;
                Properties.Settings.Default.Save();
                UpdateCheckbox();
'@
$newSelectionSave = @'
                ThisAddIn.Model = selectedModel;
                UpdateCheckbox();
'@
if ($forge.Contains($oldSelectionSave)) {
    $forge = $forge.Replace($oldSelectionSave, $newSelectionSave)
}

$oldDefaultHandler = @'
                if (this.DefaultCheckBox.Checked)
                    Properties.Settings.Default.DefaultModel = GetSelectedItemLabel();
                else
                    Properties.Settings.Default.DefaultModel = null;
'@
$newDefaultHandler = @'
                if (this.DefaultCheckBox.Checked)
                    Properties.Settings.Default.DefaultModel = GetSelectedItemLabel();
                else
                    Properties.Settings.Default.DefaultModel = null;

                Properties.Settings.Default.Save();
                UpdateCheckbox();
                SetStatus(this.DefaultCheckBox.Checked ? "★ Модель по умолчанию" : "● Готово");
'@
if ($forge.Contains($oldDefaultHandler)) {
    $forge = $forge.Replace($oldDefaultHandler, $newDefaultHandler)
}

Set-Content $forgePath $forge -Encoding UTF8

Write-Host 'RAG follow-up retrieval and explicit default-model selection prepared.'
