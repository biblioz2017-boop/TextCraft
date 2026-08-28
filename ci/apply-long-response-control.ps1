$ErrorActionPreference = 'Stop'

$script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Read-Utf8Text([string]$Path) {
    $fullPath = (Resolve-Path $Path).Path
    return [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8)
}

function Write-Utf8Text([string]$Path, [string]$Text) {
    $fullPath = (Resolve-Path $Path).Path
    [System.IO.File]::WriteAllText($fullPath, $Text, $script:Utf8NoBom)
}

Write-Host 'Applying long-response output budget control...'

$sciencePath = 'GenerateUserControl.Science.cs'
$science = Read-Utf8Text $sciencePath

if (-not $science.Contains('private ComboBox _responseLengthComboBox;')) {
    $fieldAnchor = '        private ToolTip _forceRagToolTip;'
    if (-not $science.Contains($fieldAnchor)) {
        throw 'Could not locate force-RAG UI fields for response-length setting.'
    }

    $fieldInsert = @'
        private ToolTip _forceRagToolTip;
        private Label _responseLengthLabel;
        private ComboBox _responseLengthComboBox;
        private static int _preferredMaxOutputTokens = 2600;
'@
    $science = $science.Replace($fieldAnchor, $fieldInsert.TrimEnd("`r", "`n"))
}

if (-not $science.Contains('AddResponseLengthSetting();')) {
    $loadAnchor = @'
            AddForceRagCheckbox();
            AddEvidencePanel();
'@
    $loadInsert = @'
            AddForceRagCheckbox();
            AddResponseLengthSetting();
            AddEvidencePanel();
'@
    if (-not $science.Contains($loadAnchor)) {
        throw 'Could not locate scientific pane initialization for response-length setting.'
    }
    $science = $science.Replace($loadAnchor, $loadInsert)
}

if (-not $science.Contains('private void AddResponseLengthSetting()')) {
    $methodAnchor = '        private void AddEvidencePanel()'
    if (-not $science.Contains($methodAnchor)) {
        throw 'Could not locate evidence panel method for response-length UI.'
    }

    $method = @'
        private void AddResponseLengthSetting()
        {
            _responseLengthLabel = new Label
            {
                Text = "Длина ответа:",
                AutoSize = true,
                Height = 26,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(4, 6, 2, 2)
            };

            _responseLengthComboBox = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 178,
                Height = 28,
                Margin = new Padding(2, 3, 6, 2)
            };
            _responseLengthComboBox.Items.Add("Обычный (~1400 ток.)");
            _responseLengthComboBox.Items.Add("Подробный (~2600 ток.)");
            _responseLengthComboBox.Items.Add("Очень длинный (~3800 ток.)");

            if (_preferredMaxOutputTokens >= 3500)
                _responseLengthComboBox.SelectedIndex = 2;
            else if (_preferredMaxOutputTokens >= 2200)
                _responseLengthComboBox.SelectedIndex = 1;
            else
                _responseLengthComboBox.SelectedIndex = 0;

            _responseLengthComboBox.SelectedIndexChanged += (s, e) =>
            {
                _preferredMaxOutputTokens = GetMaxOutputTokens();
            };

            if (_forceRagToolTip != null)
            {
                _forceRagToolTip.SetToolTip(
                    _responseLengthComboBox,
                    "Задает максимальную длину генерации и одновременно резервирует место в контекстном окне модели. " +
                    "Для рефератов и литературных обзоров рекомендуется «Подробный» или «Очень длинный»."
                );
            }

            _quickActionsPanel.WrapContents = true;
            _quickActionsPanel.Controls.Add(_responseLengthLabel);
            _quickActionsPanel.Controls.Add(_responseLengthComboBox);

            // The control usually wraps onto a third line in the narrow Word task pane.
            if (_mainLayout != null && _mainLayout.RowStyles.Count > 2)
                _mainLayout.RowStyles[2].Height = 96F;
        }

        private int GetMaxOutputTokens()
        {
            if (_responseLengthComboBox == null)
                return _preferredMaxOutputTokens;

            switch (_responseLengthComboBox.SelectedIndex)
            {
                case 0:
                    return 1400;
                case 2:
                    return 3800;
                default:
                    return 2600;
            }
        }

'@
    $science = $science.Replace($methodAnchor, $method + $methodAnchor)
}

if (-not $science.Contains('int maxOutputTokens = GetMaxOutputTokens();')) {
    $requestAnchor = @'
                bool forceRag = _forceRagCheckBox != null && _forceRagCheckBox.Checked;
                string retrievalQuery = BuildRagRetrievalQuery(userQuery);
'@
    $requestInsert = @'
                bool forceRag = _forceRagCheckBox != null && _forceRagCheckBox.Checked;
                int maxOutputTokens = GetMaxOutputTokens();
                string retrievalQuery = BuildRagRetrievalQuery(userQuery);
'@
    if (-not $science.Contains($requestAnchor)) {
        throw 'Could not locate RAG request settings for max-output token selection.'
    }
    $science = $science.Replace($requestAnchor, $requestInsert)
}

$oldAsk = @'
                var streamingAnswer = RAGControl.AskQuestion(
                    new SystemChatMessage(systemPrompt),
                    messages,
                    docRange,
                    GetTemperature()
                );
'@
$newAsk = @'
                var streamingAnswer = RAGControl.AskQuestion(
                    new SystemChatMessage(systemPrompt),
                    messages,
                    docRange,
                    GetTemperature(),
                    maxOutputTokens
                );
'@
if ($science.Contains($oldAsk)) {
    $science = $science.Replace($oldAsk, $newAsk)
} elseif (-not $science.Contains('                    maxOutputTokens')) {
    throw 'Could not locate RAG-aware AskQuestion call for output token budget.'
}

$science = $science.Replace(
    '                _responseLabel.Text = "Диалог / ответ — ожидаю первый токен…";',
    '                _responseLabel.Text = "Диалог / ответ — ожидаю первый токен… до " + maxOutputTokens + " ток.";'
)

Write-Utf8Text $sciencePath $science

$ragPath = 'RAGControl.cs'
$rag = Read-Utf8Text $ragPath

if (-not $rag.Contains('ProcessInformationWithOutputBudget')) {
    $imageAnchor = '        public static Task<ClientResult<GeneratedImage>> AskQuestionForImage'
    if (-not $rag.Contains($imageAnchor)) {
        throw 'Could not locate image AskQuestion method for long-response overload.'
    }

    $longAsk = @'
        public static AsyncCollectionResult<StreamingChatCompletionUpdate> AskQuestion(
            SystemChatMessage systemPrompt,
            IEnumerable<ChatMessage> messages,
            Word.Range context,
            float temperature,
            int maxOutputTokens,
            Word.Document doc = null
        )
        {
            int contextLength = Math.Max(2048, ThisAddIn.ContextLength);
            int safeMaxOutputTokens = Math.Max(
                512,
                Math.Min(maxOutputTokens, Math.Max(512, contextLength - 2048))
            );

            var chatHistory = ProcessInformationWithOutputBudget(
                systemPrompt,
                messages,
                context,
                safeMaxOutputTokens,
                doc
            );

            ChatClient client = new ChatClient(
                ThisAddIn.Model,
                new ApiKeyCredential(ThisAddIn.ApiKey),
                ThisAddIn.ClientOptions
            );

            ChatCompletionOptions options = new ChatCompletionOptions()
            {
                Temperature = temperature * 2,
                MaxOutputTokenCount = safeMaxOutputTokens
            };

            return client.CompleteChatStreamingAsync(
                chatHistory,
                options,
                ThisAddIn.CancellationTokenSource.Token
            );
        }

'@
    $rag = $rag.Replace($imageAnchor, $longAsk + $imageAnchor)

    $processAnchor = '        private static List<ChatMessage> ProcessInformation('
    if (-not $rag.Contains($processAnchor)) {
        throw 'Could not locate ProcessInformation for reserved output context.'
    }

    $budgetMethod = @'
        private static List<ChatMessage> ProcessInformationWithOutputBudget(
            SystemChatMessage systemPrompt,
            IEnumerable<ChatMessage> messages,
            Word.Range context,
            int maxOutputTokens,
            Word.Document doc = null
        )
        {
            if (doc == null)
                doc = context.Document;

            string document = context.Text ?? string.Empty;
            int userPromptLen = GetUserPromptLen(messages);
            ChatMessage lastUserPrompt = messages.Last();
            string semanticRagQuery = lastUserPrompt.Content[0].Text ?? string.Empty;

            const string ragTopicMarker = "Тема для семантического поиска в RAG:";
            int ragTopicStart = semanticRagQuery.IndexOf(ragTopicMarker, StringComparison.OrdinalIgnoreCase);
            if (ragTopicStart >= 0)
            {
                ragTopicStart += ragTopicMarker.Length;
                int ragTopicEnd = semanticRagQuery.IndexOf('\n', ragTopicStart);
                if (ragTopicEnd < 0)
                    ragTopicEnd = semanticRagQuery.Length;

                string extractedTopic = semanticRagQuery.Substring(
                    ragTopicStart,
                    ragTopicEnd - ragTopicStart
                ).Trim();
                if (!string.IsNullOrWhiteSpace(extractedTopic))
                    semanticRagQuery = extractedTopic;
            }

            bool explicitGroundedRag = false;
            foreach (ChatMessage message in messages)
            {
                if (message.Content.Count == 0)
                    continue;

                string text = message.Content[0].Text ?? string.Empty;
                if (text.IndexOf("ПРОВЕРЕННЫЕ RAG-ФРАГМЕНТЫ", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    explicitGroundedRag = true;
                    break;
                }
            }

            int contextLength = Math.Max(2048, ThisAddIn.ContextLength);
            int safetyTokens = Math.Max(256, (int)(contextLength * 0.05));
            int availablePromptTokens = Math.Max(
                1024,
                contextLength - maxOutputTokens - safetyTokens
            );
            float promptPercentage = Math.Max(
                0.30f,
                Math.Min(0.80f, (float)availablePromptTokens / (float)contextLength)
            );

            var constraints = RAGControl.OptimizeConstraint(
                promptPercentage,
                contextLength,
                CommonUtils.CharToTokenCount(systemPrompt.Content[0].Text.Length + userPromptLen),
                CommonUtils.CharToTokenCount(document.Length)
            );

            // Strict RAG already contains curated [S#] evidence and local cursor context.
            // Do not inject the same PDF chunks a second time and do not spend the output
            // reserve on the entire Word document. This is the main protection against a
            // response ending in the middle of a sentence on an 8K local context window.
            if (explicitGroundedRag)
            {
                document = string.Empty;
            }
            else if (constraints["document_content_rag"] == 1f)
            {
                document = RAGControl.GetWordDocumentAsRAG(semanticRagQuery, context);
            }

            string ragQuery = string.Empty;
            if (!explicitGroundedRag && constraints["rag_context"] != 0f)
            {
                ragQuery = ThisAddIn.AllTaskPanes[doc].Item3.GetRAGContext(
                    semanticRagQuery,
                    (int)(contextLength * constraints["rag_context"])
                );
            }

            List<ChatMessage> chatHistory = new List<ChatMessage>()
            {
                systemPrompt,
                new UserChatMessage($@"{Forge.CultureHelper.GetLocalizedString("(RAGControl.cs) [AskQuestion] chatHistory #1")}\n""{CommonUtils.SubstringTokens(document, (int)(contextLength * constraints["document_content"]))}""")
            };

            if (ragQuery != string.Empty)
            {
                chatHistory.Add(
                    new UserChatMessage($@"{Forge.CultureHelper.GetLocalizedString("(RAGControl.cs) [AskQuestion] chatHistory #2")}\n""{ragQuery}""")
                );
            }

            chatHistory.AddRange(messages);
            return chatHistory;
        }

'@
    $rag = $rag.Replace($processAnchor, $budgetMethod + $processAnchor)
}

Write-Utf8Text $ragPath $rag

Write-Host 'Long-response control prepared: explicit output token cap, reserved context, and duplicate strict-RAG injection disabled.'
