$ErrorActionPreference = 'Stop'

# This patch is invoked by Windows PowerShell 5.1 from MSBuild. GitHub source files
# are UTF-8, frequently without a BOM, while PowerShell 5.1 may otherwise read them
# using the active ANSI code page. Use .NET UTF-8 I/O explicitly so Unicode character
# literals in Forge.cs and Russian UI strings are never corrupted during the build.
$script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Read-Utf8Text([string]$Path) {
    $fullPath = (Resolve-Path $Path).Path
    return [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8)
}

function Write-Utf8Text([string]$Path, [string]$Text) {
    $fullPath = (Resolve-Path $Path).Path
    [System.IO.File]::WriteAllText($fullPath, $Text, $script:Utf8NoBom)
}

Write-Host 'Applying RAG chat semantic-query patch...'
$ragPath = 'RAGControl.cs'
$rag = Read-Utf8Text $ragPath

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

Write-Utf8Text $ragPath $rag

Write-Host 'Applying force-RAG supplement checkbox...'
$sciencePath = 'GenerateUserControl.Science.cs'
$science = Read-Utf8Text $sciencePath

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

# Keep the chat action buttons visible. The scientific evidence panel occupies the
# bottom of the task pane, so move the response action toolbar into that panel rather
# than leaving it in a row that can be clipped on narrow Word windows.
if (-not $science.Contains('_evidencePanel.Controls.Add(_responseActionsPanel);')) {
    $oldEvidencePanel = @'
            _evidencePanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 155,
                Padding = new Padding(8, 0, 8, 8)
            };
            _evidencePanel.Controls.Add(_evidenceTextBox);
            _evidencePanel.Controls.Add(evidenceLabel);

            Controls.Add(_evidencePanel);
            Controls.SetChildIndex(_evidencePanel, 0);
'@
    $newEvidencePanel = @'
            _evidencePanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 205,
                Padding = new Padding(8, 0, 8, 8)
            };

            if (_mainLayout != null && _mainLayout.RowStyles.Count > 8)
                _mainLayout.RowStyles[8].Height = 0F;

            _responseActionsPanel.Dock = DockStyle.Bottom;
            _responseActionsPanel.Height = 64;
            _responseActionsPanel.WrapContents = true;
            _responseActionsPanel.AutoScroll = true;
            _responseActionsPanel.Margin = new Padding(0, 4, 0, 0);

            _evidencePanel.Controls.Add(_evidenceTextBox);
            _evidencePanel.Controls.Add(evidenceLabel);
            _evidencePanel.Controls.Add(_responseActionsPanel);
            _responseActionsPanel.BringToFront();

            Controls.Add(_evidencePanel);
            Controls.SetChildIndex(_evidencePanel, 0);
'@
    if (-not $science.Contains($oldEvidencePanel)) {
        throw 'Could not locate evidence panel for visible chat actions.'
    }
    $science = $science.Replace($oldEvidencePanel, $newEvidencePanel)
}

Write-Utf8Text $sciencePath $science

Write-Host 'Applying chat-to-Word actions...'
$chatPath = 'GenerateUserControl.cs'
$chat = Read-Utf8Text $chatPath

if (-not $chat.Contains('private Button _insertWholeChatButton;')) {
    $chatFieldsOld = @'
        private Button _insertButton;
        private Button _copyButton;
'@
    $chatFieldsNew = @'
        private Button _insertButton;
        private Button _insertWholeChatButton;
        private Button _insertSelectedChatButton;
        private Button _copyButton;
'@
    if (-not $chat.Contains($chatFieldsOld)) {
        throw 'Could not locate chat action fields.'
    }
    $chat = $chat.Replace($chatFieldsOld, $chatFieldsNew)
}

$chat = $chat.Replace('                Text = "Вставить в документ",', '                Text = "Ответ → Word",')
$chat = $chat.Replace('                Text = "Копировать",', '                Text = "Копировать ответ",')
$chat = $chat.Replace('                Text = "Очистить",', '                Text = "Очистить чат",')

if (-not $chat.Contains('_responseTextBox.SelectionChanged += ResponseTextBox_SelectionChanged;')) {
    $actionsPanelAnchor = '            _responseActionsPanel = new FlowLayoutPanel'
    if (-not $chat.Contains($actionsPanelAnchor)) {
        throw 'Could not locate response action panel anchor.'
    }
    $chat = $chat.Replace(
        $actionsPanelAnchor,
        "            _responseTextBox.SelectionChanged += ResponseTextBox_SelectionChanged;`r`n            _responseTextBox.TextChanged += ResponseTextBox_TextChanged;`r`n`r`n" + $actionsPanelAnchor
    )
}

if (-not $chat.Contains('_insertWholeChatButton = new Button')) {
    $insertButtonAnchor = '            _insertButton.Click += InsertButton_Click;'
    $insertButtons = @'
            _insertButton.Click += InsertButton_Click;

            _insertWholeChatButton = new Button
            {
                Text = "Весь чат → Word",
                AutoSize = true,
                Height = 28,
                Enabled = false
            };
            _insertWholeChatButton.Click += InsertWholeChatButton_Click;

            _insertSelectedChatButton = new Button
            {
                Text = "Выделенное → Word",
                AutoSize = true,
                Height = 28,
                Enabled = false
            };
            _insertSelectedChatButton.Click += InsertSelectedChatButton_Click;
'@
    if (-not $chat.Contains($insertButtonAnchor)) {
        throw 'Could not locate last-response insert button.'
    }
    $chat = $chat.Replace($insertButtonAnchor, $insertButtons.TrimEnd("`r", "`n"))
}

if (-not $chat.Contains('_responseActionsPanel.Controls.Add(_insertWholeChatButton);')) {
    $oldActionList = @'
            _responseActionsPanel.Controls.Add(_insertButton);
            _responseActionsPanel.Controls.Add(_copyButton);
            _responseActionsPanel.Controls.Add(_clearButton);
'@
    $newActionList = @'
            _responseActionsPanel.Controls.Add(_insertButton);
            _responseActionsPanel.Controls.Add(_insertWholeChatButton);
            _responseActionsPanel.Controls.Add(_insertSelectedChatButton);
            _responseActionsPanel.Controls.Add(_copyButton);
            _responseActionsPanel.Controls.Add(_clearButton);
'@
    if (-not $chat.Contains($oldActionList)) {
        throw 'Could not locate response action control list.'
    }
    $chat = $chat.Replace($oldActionList, $newActionList)
}

if (-not $chat.Contains('private void InsertWholeChatButton_Click')) {
    $copyHandlerAnchor = '        private void CopyButton_Click(object sender, EventArgs e)'
    $chatHandlers = @'
        private void InsertWholeChatButton_Click(object sender, EventArgs e)
        {
            try
            {
                string chatText = (_responseTextBox.Text ?? string.Empty).Trim();
                if (chatText.Length == 0)
                    return;

                InsertPlainChatTextIntoDocument(chatText);
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void InsertSelectedChatButton_Click(object sender, EventArgs e)
        {
            try
            {
                string selectedText = _responseTextBox.SelectedText ?? string.Empty;
                if (string.IsNullOrWhiteSpace(selectedText))
                {
                    MessageBox.Show(
                        "Сначала выделите мышью нужный фрагмент в окне «Диалог / ответ».",
                        "TextCraft",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    return;
                }

                InsertPlainChatTextIntoDocument(selectedText);
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private static void InsertPlainChatTextIntoDocument(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return;

            Word.Range insertionRange = Globals.ThisAddIn.Application.Selection.Range.Duplicate;
            insertionRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            insertionRange.Text = WordMarkdown.RemoveMarkdownSyntax(text.Trim());
            Globals.ThisAddIn.Application.Selection.SetRange(insertionRange.End, insertionRange.End);
        }

        private void ResponseTextBox_SelectionChanged(object sender, EventArgs e)
        {
            UpdateChatActionButtons();
        }

        private void ResponseTextBox_TextChanged(object sender, EventArgs e)
        {
            UpdateChatActionButtons();
        }

        private void UpdateChatActionButtons()
        {
            if (_responseTextBox == null)
                return;

            bool hasChat = _responseTextBox.TextLength > 0;
            if (_insertWholeChatButton != null)
                _insertWholeChatButton.Enabled = hasChat;
            if (_insertSelectedChatButton != null)
                _insertSelectedChatButton.Enabled = hasChat && _responseTextBox.SelectionLength > 0;
            if (_clearButton != null)
                _clearButton.Enabled = hasChat;
        }

'@
    if (-not $chat.Contains($copyHandlerAnchor)) {
        throw 'Could not locate CopyButton handler for chat action methods.'
    }
    $chat = $chat.Replace($copyHandlerAnchor, $chatHandlers + $copyHandlerAnchor)
}

$oldClearTail = @'
            _lastTemplateName = string.Empty;
            _insertButton.Enabled = false;
            _copyButton.Enabled = false;
        }
'@
$newClearTail = @'
            _lastTemplateName = string.Empty;
            _insertButton.Enabled = false;
            _copyButton.Enabled = false;
            UpdateChatActionButtons();
        }
'@
if ($chat.Contains($oldClearTail)) {
    $chat = $chat.Replace($oldClearTail, $newClearTail)
}

Write-Utf8Text $chatPath $chat

Write-Host 'Applying explicit default-model UI patch...'
$designerPath = 'Forge.Designer.cs'
$designer = Read-Utf8Text $designerPath

# Restore the upstream checkbox as an explicit user action in the Model group.
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
Write-Utf8Text $designerPath $designer

$forgePath = 'Forge.cs'
$forge = Read-Utf8Text $forgePath

# Selecting a model changes only the current session. It becomes the startup model
# only after the user explicitly checks the default-model checkbox.
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

Write-Utf8Text $forgePath $forge

Write-Host 'RAG follow-up retrieval, strict chat actions, UTF-8-safe patching, and explicit default-model selection prepared.'
