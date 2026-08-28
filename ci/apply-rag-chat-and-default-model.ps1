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

Write-Host 'Applying force-RAG supplement checkbox...'
$sciencePath = 'GenerateUserControl.Science.cs'
$science = Get-Content $sciencePath -Raw

if (-not $science.Contains('private CheckBox _forceRagCheckBox;')) {
    $fieldAnchor = '        private Timer _chatRedrawTimer;'
    if (-not $science.Contains($fieldAnchor)) {
        throw 'Could not locate chat UI field anchor for force-RAG checkbox.'
    }

    $fieldInsert = @'
        private Timer _chatRedrawTimer;
        private CheckBox _forceRagCheckBox;
        private ToolTip _forceRagToolTip;
'@
    $science = $science.Replace($fieldAnchor, $fieldInsert.TrimEnd("`r", "`n"))
}

if (-not $science.Contains('AddForceRagCheckbox();')) {
    $loadAnchor = @'
            AddScientificQuickActions();
            AddEvidencePanel();
'@
    $loadInsert = @'
            AddScientificQuickActions();
            AddForceRagCheckbox();
            AddEvidencePanel();
'@
    if (-not $science.Contains($loadAnchor)) {
        throw 'Could not locate chat OnLoad action block for force-RAG checkbox.'
    }
    $science = $science.Replace($loadAnchor, $loadInsert)
}

if (-not $science.Contains('private void AddForceRagCheckbox()')) {
    $methodAnchor = '        private void AddEvidencePanel()'
    $methodInsert = @'
        private void AddForceRagCheckbox()
        {
            _forceRagCheckBox = new CheckBox
            {
                Text = "RAG: использовать и дополнять текст",
                Checked = true,
                AutoSize = true,
                Height = 26,
                Margin = new Padding(2, 3, 6, 2)
            };

            _forceRagToolTip = new ToolTip();
            _forceRagToolTip.SetToolTip(
                _forceRagCheckBox,
                "Принудительно использовать отмеченные PDF/RAG и дополнять ответ новыми подтвержденными сведениями. " +
                "Можно оставить включенным и не писать каждый раз 'используй материал из RAG'."
            );

            // The quick-action row was originally only 34 px high. Allow a second line so
            // the checkbox remains visible even in a narrow Word task pane.
            _quickActionsPanel.WrapContents = true;
            if (_mainLayout != null && _mainLayout.RowStyles.Count > 2)
                _mainLayout.RowStyles[2].Height = 62F;

            _quickActionsPanel.Controls.Add(_forceRagCheckBox);
            _quickActionsPanel.Controls.SetChildIndex(_forceRagCheckBox, 0);
        }

'@
    if (-not $science.Contains($methodAnchor)) {
        throw 'Could not locate evidence-panel method anchor for force-RAG checkbox.'
    }
    $science = $science.Replace($methodAnchor, $methodInsert + $methodAnchor)
}

$oldAlways = '            messages.Add(new UserChatMessage(AlwaysUseRagInstruction));'
$newAlways = @'
            if (_forceRagCheckBox != null && _forceRagCheckBox.Checked)
                messages.Add(new UserChatMessage(AlwaysUseRagInstruction));
'@
if ($science.Contains($oldAlways)) {
    $science = $science.Replace($oldAlways, $newAlways.TrimEnd("`r", "`n"))
}

$oldFinalPrompt = @'
            string retrievalQuery = BuildRagRetrievalQuery(userQuery);
            string finalPrompt =
                "Текущий запрос пользователя: " + userQuery + "\n" +
                "Тема для семантического поиска в RAG: " + retrievalQuery + "\n\n" +
                "Ответь именно на текущий запрос. Если выше передан RAG-контекст, обязательно используй его. " +
                "Для продолжения предыдущего ответа добавляй новые сведения из RAG и не повторяй старый текст дословно.";
'@
$newFinalPrompt = @'
            string retrievalQuery = BuildRagRetrievalQuery(userQuery);
            bool forceRag = _forceRagCheckBox != null && _forceRagCheckBox.Checked;
            string ragBehavior = forceRag
                ? "Режим принудительного RAG включен. Обязательно используй релевантные сведения из отмеченных PDF, " +
                  "дополни ими текст или предыдущий ответ, добавляй только новые подтвержденные сведения и сохраняй ссылки на источник/страницу. " +
                  "Не повторяй предыдущий ответ целиком и не подменяй RAG знаниями модели."
                : "Ответь на текущий запрос с учетом предыдущего диалога. RAG-контекст можно использовать как дополнительный источник, если он релевантен.";

            string finalPrompt =
                "Текущий запрос пользователя: " + userQuery + "\n" +
                "Тема для семантического поиска в RAG: " + retrievalQuery + "\n\n" +
                ragBehavior;
'@
if ($science.Contains($oldFinalPrompt)) {
    $science = $science.Replace($oldFinalPrompt, $newFinalPrompt)
}

Set-Content $sciencePath $science -Encoding UTF8

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

Write-Host 'RAG follow-up retrieval, force-RAG checkbox, and explicit default-model selection prepared.'
