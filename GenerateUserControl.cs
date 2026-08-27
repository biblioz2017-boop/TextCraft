using System;
using System.ClientModel;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenAI.Chat;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    public partial class GenerateUserControl : UserControl
    {
        public static readonly CultureLocalizationHelper CultureHelper =
            new CultureLocalizationHelper("TextForge.GenerateUserControl", typeof(GenerateUserControl).Assembly);

        private const int DefaultLocalContextBefore = 2;
        private const int DefaultLocalContextAfter = 2;
        private const int MaxLocalContextParagraphs = 20;
        private const int MaxConversationTurns = 3;

        private ComboBox _ragTemplateComboBox;
        private Label _ragTemplateLabel;
        private Label _queryLabel;
        private Label _responseLabel;
        private TableLayoutPanel _mainLayout;
        private FlowLayoutPanel _quickActionsPanel;
        private FlowLayoutPanel _responseActionsPanel;
        private RichTextBox _responseTextBox;
        private Button _checkSelectionButton;
        private Button _matrixButton;
        private Button _insertButton;
        private Button _copyButton;
        private Button _clearButton;
        private List<RagPromptTemplate> _ragTemplates;
        private readonly List<ConversationTurn> _conversationTurns = new List<ConversationTurn>();

        private string _lastResponseMarkdown = string.Empty;
        private string _lastTemplateName = string.Empty;

        private const string RagEvidenceRules =
            "Работай как научный аналитик по предоставленным материалам. " +
            "Основой ответа должны быть фрагменты из подключенных PDF/RAG и, при необходимости, текущий Word-документ. " +
            "Не заполняй пробелы внешними знаниями и не выдумывай сведения. " +
            "Если материалов недостаточно для утверждения, прямо укажи это. " +
            "Не выдумывай авторов, названия работ, годы, DOI, номера страниц или библиографические данные. " +
            "Если в RAG-контексте присутствуют метки вида [Source: имя.pdf; Page: N], используй их как источник и страницу. " +
            "После ключевых фактических положений указывай ссылку в формате [имя.pdf, с. N], только если такая страница явно дана в контексте. " +
            "Если несколько источников расходятся, не сглаживай расхождения: покажи позиции отдельно. ";

        public GenerateUserControl()
        {
            try
            {
                InitializeComponent();
                MatchScrollBarTemperature();
                InitializeRagTemplates();
                ConfigureChatLayout();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void InitializeRagTemplates()
        {
            _ragTemplates = new List<RagPromptTemplate>
            {
                new RagPromptTemplate(
                    "Свободный запрос",
                    string.Empty
                ),
                new RagPromptTemplate(
                    "Литературный обзор",
                    RagEvidenceRules +
                    "Подготовь структурированный литературный обзор по теме, указанной пользователем. " +
                    "Синтезируй литературу по проблемам, подходам и выводам, а не пересказывай PDF по очереди. " +
                    "Покажи постановку проблемы, основные направления исследований, согласующиеся результаты, противоречия, ограничения, пробелы и итоговый вывод. " +
                    "Сохраняй строгий научный стиль. Отделяй подтвержденные материалами выводы от осторожных интерпретаций."
                ),
                new RagPromptTemplate(
                    "Раздел литобзора",
                    RagEvidenceRules +
                    "Напиши связный черновик раздела литературного обзора по теме пользователя. " +
                    "Это должен быть цельный научный текст с логическими переходами между абзацами, а не список аннотаций. " +
                    "Объединяй близкие результаты разных источников, отдельно отмечай несогласующиеся данные и методические ограничения. " +
                    "Не добавляй фактов, которых нет в переданных материалах."
                ),
                new RagPromptTemplate(
                    "Аналитическая справка",
                    RagEvidenceRules +
                    "Составь аналитическую справку по теме или вопросу пользователя. " +
                    "Структура: краткий ответ; ключевые факты; позиции и результаты источников; спорные или неопределенные моменты; практический или научный вывод. " +
                    "Пиши компактно, но содержательно."
                ),
                new RagPromptTemplate(
                    "Сравнение источников",
                    RagEvidenceRules +
                    "Сравни подключенные источники по вопросу пользователя. " +
                    "Сначала дай краткий синтез, затем сравни объект или контекст, метод или подход, основные результаты, ограничения и различия между источниками. " +
                    "Если какой-либо параметр в найденных фрагментах отсутствует, пиши 'нет данных', а не предполагай его."
                ),
                new RagPromptTemplate(
                    "Противоречия в литературе",
                    RagEvidenceRules +
                    "Найди и систематизируй противоречия, несовпадения и разные интерпретации по теме пользователя. " +
                    "Для каждого противоречия укажи позицию или результат A, позицию или результат B, источники, условия или методические различия, если они прямо описаны, и что остается неясным. " +
                    "Не придумывай причины противоречий, если материалы их не подтверждают."
                ),
                new RagPromptTemplate(
                    "Матрица литературы",
                    RagEvidenceRules +
                    "Составь исследовательскую матрицу по теме пользователя. " +
                    "Верни только таблицу без вводного и заключительного текста. " +
                    "Используй одну строку на источник и следующие столбцы: Источник; Цель; Объект/выборка; Метод; Показатели/переменные; Основные результаты; Ограничения; Релевантность теме. " +
                    "Для отсутствующих сведений пиши 'нет данных'. " +
                    "Предпочтительный формат ответа — строки с полями, разделенными символами табуляции. Первая строка должна содержать названия столбцов."
                ),
                new RagPromptTemplate(
                    "Пробелы и перспективы",
                    RagEvidenceRules +
                    "Определи исследовательские пробелы и перспективные направления по теме пользователя. " +
                    "Сначала перечисли пробелы, прямо обозначенные авторами или вытекающие из ограничений и противоречий в переданных материалах. " +
                    "Для каждого пункта отметь, является ли он прямо указанным в источнике или аналитическим выводом на основе нескольких источников. " +
                    "Не предлагай направления, не связанные с содержанием материалов."
                ),
                new RagPromptTemplate(
                    "Проверка тезисов",
                    RagEvidenceRules +
                    "Разбей текст пользователя на отдельные проверяемые тезисы. " +
                    "Для каждого тезиса определи один статус: подтверждается; частично подтверждается; противоречит источникам; подтверждение не найдено. " +
                    "Покажи краткое доказательство из найденных материалов, источник и страницу, если страница дана в RAG-контексте. " +
                    "Если тезис содержит несколько утверждений с разной доказательной базой, раздели его. " +
                    "Не считать отсутствие найденного фрагмента доказательством ложности тезиса. " +
                    "В конце дай короткий список тезисов, для которых нужно подобрать дополнительные источники."
                ),
                new RagPromptTemplate(
                    "Конспект источников",
                    RagEvidenceRules +
                    "Сделай тематический конспект материалов по запросу пользователя. " +
                    "Сгруппируй сведения по смысловым блокам. Для каждого блока перечисли основные положения и источники, где они встречаются. " +
                    "Не дублируй одинаковые утверждения; отмечай, если одно положение подтверждается несколькими PDF."
                )
            };
        }

        private void ConfigureChatLayout()
        {
            _ragTemplateLabel = new Label
            {
                Text = "Шаблон для литературы / RAG:",
                AutoSize = true,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 0, 0, 3)
            };

            _ragTemplateComboBox = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 6)
            };

            foreach (var template in _ragTemplates)
                _ragTemplateComboBox.Items.Add(template);

            _ragTemplateComboBox.SelectedIndex = 0;
            _ragTemplateComboBox.SelectedIndexChanged += RagTemplateComboBox_SelectedIndexChanged;

            _quickActionsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0, 0, 0, 6)
            };

            _checkSelectionButton = new Button
            {
                Text = "Проверить выделенное",
                AutoSize = true,
                Height = 28
            };
            _checkSelectionButton.Click += CheckSelectionButton_Click;

            _matrixButton = new Button
            {
                Text = "Матрица",
                AutoSize = true,
                Height = 28
            };
            _matrixButton.Click += MatrixButton_Click;

            _quickActionsPanel.Controls.Add(_checkSelectionButton);
            _quickActionsPanel.Controls.Add(_matrixButton);

            _queryLabel = new Label
            {
                Text = "Запрос:",
                AutoSize = true,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 0, 0, 3)
            };

            PromptTextBox.Multiline = true;
            PromptTextBox.ScrollBars = ScrollBars.Vertical;
            PromptTextBox.AcceptsReturn = true;
            PromptTextBox.Dock = DockStyle.Fill;
            PromptTextBox.Margin = new Padding(0, 0, 0, 6);

            GenerateButton.Text = "Спросить";
            GenerateButton.Dock = DockStyle.Fill;
            GenerateButton.Margin = new Padding(0, 0, 0, 6);

            _responseLabel = new Label
            {
                Text = "Диалог / ответ:",
                AutoSize = true,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 0, 0, 3)
            };

            _responseTextBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                DetectUrls = true,
                HideSelection = false,
                BackColor = SystemColors.Window,
                Margin = new Padding(0, 0, 0, 6)
            };

            _responseActionsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0)
            };

            _insertButton = new Button
            {
                Text = "Вставить в документ",
                AutoSize = true,
                Height = 28,
                Enabled = false
            };
            _insertButton.Click += InsertButton_Click;

            _copyButton = new Button
            {
                Text = "Копировать",
                AutoSize = true,
                Height = 28,
                Enabled = false
            };
            _copyButton.Click += CopyButton_Click;

            _clearButton = new Button
            {
                Text = "Очистить",
                AutoSize = true,
                Height = 28
            };
            _clearButton.Click += ClearButton_Click;

            _responseActionsPanel.Controls.Add(_insertButton);
            _responseActionsPanel.Controls.Add(_copyButton);
            _responseActionsPanel.Controls.Add(_clearButton);

            _mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 9,
                Padding = new Padding(8)
            };
            _mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            _mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 22F));
            _mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));
            _mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));
            _mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 22F));
            _mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 90F));
            _mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38F));
            _mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 22F));
            _mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            _mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36F));

            this.Controls.Remove(PromptTextBox);
            this.Controls.Remove(GenerateButton);

            _mainLayout.Controls.Add(_ragTemplateLabel, 0, 0);
            _mainLayout.Controls.Add(_ragTemplateComboBox, 0, 1);
            _mainLayout.Controls.Add(_quickActionsPanel, 0, 2);
            _mainLayout.Controls.Add(_queryLabel, 0, 3);
            _mainLayout.Controls.Add(PromptTextBox, 0, 4);
            _mainLayout.Controls.Add(GenerateButton, 0, 5);
            _mainLayout.Controls.Add(_responseLabel, 0, 6);
            _mainLayout.Controls.Add(_responseTextBox, 0, 7);
            _mainLayout.Controls.Add(_responseActionsPanel, 0, 8);

            this.Controls.Add(_mainLayout);
            _mainLayout.BringToFront();
            panel1.Visible = false;
        }

        private void RagTemplateComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            var template = _ragTemplateComboBox.SelectedItem as RagPromptTemplate;
            if (template == null)
                return;

            if (template.Name == "Свободный запрос")
                _queryLabel.Text = "Запрос:";
            else if (template.Name == "Проверка тезисов")
                _queryLabel.Text = "Тезисы для проверки:";
            else
                _queryLabel.Text = "Тема / вопрос:";
        }

        private void CheckSelectionButton_Click(object sender, EventArgs e)
        {
            try
            {
                Word.Selection selection = Globals.ThisAddIn.Application.Selection;
                if (selection == null || selection.End <= selection.Start)
                {
                    MessageBox.Show(
                        "Выделите в Word абзац или несколько тезисов, которые нужно проверить по подключенной литературе.",
                        "TextCraft",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    return;
                }

                SelectTemplate("Проверка тезисов");
                PromptTextBox.Text = (selection.Text ?? string.Empty).Trim();
                GenerateButton.PerformClick();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void MatrixButton_Click(object sender, EventArgs e)
        {
            try
            {
                SelectTemplate("Матрица литературы");

                Word.Selection selection = Globals.ThisAddIn.Application.Selection;
                if (selection != null && selection.End > selection.Start)
                {
                    string selectedText = (selection.Text ?? string.Empty).Trim();
                    if (selectedText.Length > 0 && selectedText.Length <= 1200)
                        PromptTextBox.Text = selectedText;
                }

                PromptTextBox.Focus();
                PromptTextBox.SelectionStart = PromptTextBox.TextLength;
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void SelectTemplate(string templateName)
        {
            for (int i = 0; i < _ragTemplateComboBox.Items.Count; i++)
            {
                var template = _ragTemplateComboBox.Items[i] as RagPromptTemplate;
                if (template != null && template.Name == templateName)
                {
                    _ragTemplateComboBox.SelectedIndex = i;
                    return;
                }
            }
        }

        private string GetSelectedTemplateInstruction()
        {
            var template = _ragTemplateComboBox.SelectedItem as RagPromptTemplate;
            return template == null ? string.Empty : template.Instruction;
        }

        private string GetSelectedTemplateName()
        {
            var template = _ragTemplateComboBox.SelectedItem as RagPromptTemplate;
            return template == null ? string.Empty : template.Name;
        }

        private async void GenerateButton_Click(object sender, EventArgs e)
        {
            try
            {
                string textBoxContent = PromptTextBox.Text.Trim();
                if (textBoxContent.Length == 0)
                    throw new EmptyTextBoxException(
                        CultureHelper.GetLocalizedString("[GenerateButton_Click] TextBoxEmptyException #1")
                    );

                if (ModelProperties.IsImageModel(ThisAddIn.Model))
                {
                    MessageBox.Show(
                        "Для чата с документом и литературой выберите языковую модель.",
                        "TextCraft",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    return;
                }

                GenerateButton.Enabled = false;
                _insertButton.Enabled = false;
                _copyButton.Enabled = false;

                Word.Range anchorRange = Globals.ThisAddIn.Application.Selection.Range.Duplicate;
                Word.Range docRange = Globals.ThisAddIn.Application.ActiveDocument.Range();
                string localCursorContext = GetLocalCursorContext(anchorRange);
                string templateInstruction = GetSelectedTemplateInstruction();
                string templateName = GetSelectedTemplateName();

                List<ChatMessage> messages = BuildMessages(
                    textBoxContent,
                    templateInstruction,
                    localCursorContext
                );

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
                _lastResponseMarkdown = response;
                _lastTemplateName = templateName;

                if (!string.IsNullOrWhiteSpace(response))
                {
                    _conversationTurns.Add(new ConversationTurn(textBoxContent, response));
                    while (_conversationTurns.Count > MaxConversationTurns)
                        _conversationTurns.RemoveAt(0);

                    _insertButton.Enabled = true;
                    _copyButton.Enabled = true;
                }

                PromptTextBox.Clear();
                PromptTextBox.Focus();
            }
            catch (EmptyTextBoxException ex)
            {
                CommonUtils.DisplayInformation(ex);
            }
            catch (OperationCanceledException ex)
            {
                CommonUtils.DisplayWarning(ex);
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
            finally
            {
                GenerateButton.Enabled = true;
            }
        }

        private List<ChatMessage> BuildMessages(
            string userQuery,
            string templateInstruction,
            string localCursorContext
        )
        {
            var messages = new List<ChatMessage>();

            if (_conversationTurns.Count > 0)
            {
                StringBuilder dialogue = new StringBuilder();
                dialogue.AppendLine("Предыдущий диалог. Используй его только для понимания уточняющих вопросов:");
                foreach (ConversationTurn turn in _conversationTurns)
                {
                    dialogue.AppendLine("Пользователь: " + turn.Question);
                    dialogue.AppendLine("Ассистент: " + turn.Answer);
                    dialogue.AppendLine();
                }

                int dialogueTokens = Math.Max(384, (int)(ThisAddIn.ContextLength * 0.15));
                messages.Add(
                    new UserChatMessage(
                        CommonUtils.SubstringTokens(dialogue.ToString(), dialogueTokens)
                    )
                );
            }

            if (!string.IsNullOrWhiteSpace(localCursorContext))
            {
                messages.Add(
                    new UserChatMessage(
                        "Локальный контекст вокруг курсора из текущего Word-документа. " +
                        "Используй его для связи ответа с текущим разделом, но не считай его внешним доказательством:\n\n\"" +
                        localCursorContext +
                        "\""
                    )
                );
            }

            if (!string.IsNullOrWhiteSpace(templateInstruction))
                messages.Add(new UserChatMessage(templateInstruction));

            // This MUST stay last: RAGControl.ProcessInformation() uses messages.Last()
            // as the semantic retrieval query for both the current document and PDF RAG.
            messages.Add(new UserChatMessage(userQuery));

            return messages;
        }

        private void AppendConversationHeader(string question, string templateName)
        {
            if (_responseTextBox.TextLength > 0)
                _responseTextBox.AppendText(Environment.NewLine + Environment.NewLine);

            _responseTextBox.AppendText("Вы");
            if (!string.IsNullOrWhiteSpace(templateName) && templateName != "Свободный запрос")
                _responseTextBox.AppendText(" [" + templateName + "]");
            _responseTextBox.AppendText(": " + question + Environment.NewLine + Environment.NewLine);
            _responseTextBox.AppendText("TextCraft:" + Environment.NewLine);
            _responseTextBox.SelectionStart = _responseTextBox.TextLength;
            _responseTextBox.ScrollToCaret();
        }

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

        private void InsertButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(_lastResponseMarkdown))
                    return;

                Word.Range insertionRange = Globals.ThisAddIn.Application.Selection.Range.Duplicate;
                insertionRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                if (_lastTemplateName == "Матрица литературы" && TryInsertMatrix(insertionRange, _lastResponseMarkdown))
                    return;

                string rawMarkdown = _lastResponseMarkdown;
                insertionRange.Text = WordMarkdown.RemoveMarkdownSyntax(rawMarkdown);

                try
                {
                    WordMarkdown.ApplyAllMarkdownFormatting(insertionRange, rawMarkdown);
                }
                catch (Exception ex)
                {
                    // Keep the already inserted plain text even if optional Markdown
                    // formatting encounters an unsupported construct.
                    CommonUtils.DisplayWarning(ex);
                }

                Globals.ThisAddIn.Application.Selection.SetRange(
                    insertionRange.End,
                    insertionRange.End
                );
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private static bool TryInsertMatrix(Word.Range insertionRange, string response)
        {
            List<string[]> rows = ParseMatrixRows(response);
            if (rows.Count < 2)
                return false;

            int columnCount = 0;
            foreach (string[] row in rows)
                columnCount = Math.Max(columnCount, row.Length);

            if (columnCount < 2)
                return false;

            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Table table = document.Tables.Add(insertionRange, rows.Count, columnCount);

            for (int r = 0; r < rows.Count; r++)
            {
                for (int c = 0; c < columnCount; c++)
                {
                    string value = c < rows[r].Length ? rows[r][c].Trim() : string.Empty;
                    table.Cell(r + 1, c + 1).Range.Text = value;
                }
            }

            table.Rows[1].Range.Bold = 1;
            table.Borders.Enable = 1;
            table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);

            Globals.ThisAddIn.Application.Selection.SetRange(table.Range.End, table.Range.End);
            return true;
        }

        private static List<string[]> ParseMatrixRows(string response)
        {
            var rows = new List<string[]>();
            string cleaned = response
                .Replace("```tsv", string.Empty)
                .Replace("```text", string.Empty)
                .Replace("```", string.Empty)
                .Trim();

            string[] lines = cleaned.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string rawLine in lines)
            {
                string line = rawLine.Trim();
                if (line.Length == 0)
                    continue;

                string[] cells;
                if (line.IndexOf('\t') >= 0)
                {
                    cells = line.Split('\t');
                }
                else if (line.StartsWith("|") && line.EndsWith("|"))
                {
                    cells = line.Trim('|').Split('|');
                }
                else
                {
                    continue;
                }

                if (IsMarkdownSeparatorRow(cells))
                    continue;

                rows.Add(cells);
            }

            return rows;
        }

        private static bool IsMarkdownSeparatorRow(string[] cells)
        {
            if (cells.Length == 0)
                return false;

            foreach (string cell in cells)
            {
                string value = cell.Trim().Replace("-", string.Empty).Replace(":", string.Empty);
                if (value.Length > 0)
                    return false;
            }
            return true;
        }

        private void CopyButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(_lastResponseMarkdown))
                    Clipboard.SetText(WordMarkdown.RemoveMarkdownSyntax(_lastResponseMarkdown));
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void ClearButton_Click(object sender, EventArgs e)
        {
            _responseTextBox.Clear();
            _conversationTurns.Clear();
            _lastResponseMarkdown = string.Empty;
            _lastTemplateName = string.Empty;
            _insertButton.Enabled = false;
            _copyButton.Enabled = false;
        }

        private static string GetLocalCursorContext(Word.Range anchorRange)
        {
            int paragraphsBefore = GetLocalContextSetting(
                "TEXTCRAFT_LOCAL_CONTEXT_BEFORE",
                DefaultLocalContextBefore
            );
            int paragraphsAfter = GetLocalContextSetting(
                "TEXTCRAFT_LOCAL_CONTEXT_AFTER",
                DefaultLocalContextAfter
            );

            Word.Range localRange = anchorRange.Duplicate;
            localRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            localRange.Expand(Word.WdUnits.wdParagraph);

            if (paragraphsBefore > 0)
                localRange.MoveStart(Word.WdUnits.wdParagraph, -paragraphsBefore);
            if (paragraphsAfter > 0)
                localRange.MoveEnd(Word.WdUnits.wdParagraph, paragraphsAfter);

            int maxTokens = Math.Max(512, (int)(ThisAddIn.ContextLength * 0.20));
            return CommonUtils.SubstringTokens(localRange.Text, maxTokens);
        }

        private static int GetLocalContextSetting(string variableName, int defaultValue)
        {
            string value =
                Environment.GetEnvironmentVariable(variableName, EnvironmentVariableTarget.User)
                ?? Environment.GetEnvironmentVariable(variableName);

            if (int.TryParse(value, out int parsed))
                return Math.Max(0, Math.Min(parsed, MaxLocalContextParagraphs));

            return defaultValue;
        }

        private void PromptTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control && e.KeyCode == Keys.Enter)
                {
                    e.SuppressKeyPress = true;
                    GenerateButton.PerformClick();
                }
                else if (e.Control && e.KeyCode == Keys.Back)
                {
                    e.SuppressKeyPress = true;
                    int cursorPosition = PromptTextBox.SelectionStart;
                    string text = PromptTextBox.Text;

                    while (cursorPosition > 0 && text[cursorPosition - 1] == ' ')
                        cursorPosition--;

                    text = text.TrimEnd();

                    if (string.IsNullOrWhiteSpace(text))
                    {
                        PromptTextBox.Clear();
                        PromptTextBox.SelectionStart = 0;
                    }
                    else
                    {
                        int lastSpaceIndex = text.LastIndexOf(' ', Math.Max(0, cursorPosition - 1));
                        if (lastSpaceIndex != -1)
                        {
                            PromptTextBox.Text = text.Remove(
                                lastSpaceIndex + 1,
                                cursorPosition - lastSpaceIndex - 1
                            );
                            PromptTextBox.SelectionStart = lastSpaceIndex + 1;
                        }
                        else
                        {
                            PromptTextBox.Text = text.Remove(0, cursorPosition);
                            PromptTextBox.SelectionStart = 0;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void TemperatureTrackBar_Scroll(object sender, EventArgs e)
        {
            try
            {
                MatchScrollBarTemperature();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void MatchScrollBarTemperature()
        {
            TemperatureValueLabel.Text = GetTemperature().ToString(
                "0.0",
                Thread.CurrentThread.CurrentUICulture
            );
        }

        private float GetTemperature()
        {
            return TemperatureTrackBar.Value / 10f;
        }

        private sealed class RagPromptTemplate
        {
            public string Name { get; private set; }
            public string Instruction { get; private set; }

            public RagPromptTemplate(string name, string instruction)
            {
                Name = name;
                Instruction = instruction;
            }

            public override string ToString()
            {
                return Name;
            }
        }

        private sealed class ConversationTurn
        {
            public string Question { get; private set; }
            public string Answer { get; private set; }

            public ConversationTurn(string question, string answer)
            {
                Question = question;
                Answer = answer;
            }
        }
    }

    public class EmptyTextBoxException : ArgumentException
    {
        public EmptyTextBoxException(string message) : base(message) { }
    }
}
