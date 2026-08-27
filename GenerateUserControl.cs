using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
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

        private ComboBox _ragTemplateComboBox;
        private Label _ragTemplateLabel;
        private Label _queryLabel;
        private TableLayoutPanel _mainLayout;
        private List<RagPromptTemplate> _ragTemplates;

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
                ConfigureSimplifiedLayout();
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
                    "Покажи: постановку проблемы; основные направления исследований; согласующиеся результаты; противоречия; ограничения; пробелы; итоговый вывод. " +
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
                    "Сначала дай краткий синтез, затем таблицу: источник; объект/контекст; метод или подход; основные результаты; ограничения; чем отличается от других источников. " +
                    "Если какой-либо параметр в найденных фрагментах отсутствует, ставь 'нет данных', а не предполагай его."
                ),
                new RagPromptTemplate(
                    "Противоречия в литературе",
                    RagEvidenceRules +
                    "Найди и систематизируй противоречия, несовпадения и разные интерпретации по теме пользователя. " +
                    "Для каждого противоречия укажи: позиция/результат A; позиция/результат B; источники; условия или методические различия, если они прямо описаны; что остается неясным. " +
                    "Не придумывай причины противоречий, если материалы их не подтверждают."
                ),
                new RagPromptTemplate(
                    "Таблица исследований",
                    RagEvidenceRules +
                    "Составь исследовательскую матрицу по теме пользователя. " +
                    "Представь таблицу со столбцами: источник; цель; объект/выборка; метод; ключевые переменные или показатели; основные результаты; ограничения; релевантность теме. " +
                    "Заполняй только то, что действительно присутствует в найденных фрагментах; для отсутствующих сведений указывай 'нет данных'."
                ),
                new RagPromptTemplate(
                    "Пробелы и перспективы",
                    RagEvidenceRules +
                    "Определи исследовательские пробелы и перспективные направления по теме пользователя. " +
                    "Сначала перечисли пробелы, прямо обозначенные авторами или очевидно вытекающие из ограничений и противоречий в переданных материалах. " +
                    "Для каждого пункта отметь, является ли он прямо указанным в источнике или аналитическим выводом на основе нескольких источников. " +
                    "Не предлагай направления, не связанные с содержанием материалов."
                ),
                new RagPromptTemplate(
                    "Проверка тезиса",
                    RagEvidenceRules +
                    "Оцени тезис пользователя по подключенным источникам. " +
                    "Раздели найденные материалы на: подтверждает тезис; частично подтверждает; противоречит; данных недостаточно. " +
                    "Для каждого пункта дай короткое обоснование и источник. В конце сформулируй осторожный итог без категоричности, если доказательств недостаточно."
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

        private void ConfigureSimplifiedLayout()
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
                Margin = new Padding(0, 0, 0, 8)
            };

            foreach (var template in _ragTemplates)
                _ragTemplateComboBox.Items.Add(template);

            _ragTemplateComboBox.SelectedIndex = 0;
            _ragTemplateComboBox.SelectedIndexChanged += RagTemplateComboBox_SelectedIndexChanged;

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
            PromptTextBox.Margin = new Padding(0, 0, 0, 8);

            GenerateButton.Text = "Выполнить";
            GenerateButton.Dock = DockStyle.Fill;
            GenerateButton.Margin = new Padding(0);

            _mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                Padding = new Padding(8)
            };
            _mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            _mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 22F));
            _mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));
            _mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 22F));
            _mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            _mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36F));

            this.Controls.Remove(PromptTextBox);
            this.Controls.Remove(GenerateButton);

            _mainLayout.Controls.Add(_ragTemplateLabel, 0, 0);
            _mainLayout.Controls.Add(_ragTemplateComboBox, 0, 1);
            _mainLayout.Controls.Add(_queryLabel, 0, 2);
            _mainLayout.Controls.Add(PromptTextBox, 0, 3);
            _mainLayout.Controls.Add(GenerateButton, 0, 4);

            this.Controls.Add(_mainLayout);
            _mainLayout.BringToFront();
            panel1.Visible = false;
        }

        private void RagTemplateComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_ragTemplateComboBox.SelectedItem is RagPromptTemplate template)
            {
                if (template.Name == "Свободный запрос")
                    _queryLabel.Text = "Запрос:";
                else if (template.Name == "Проверка тезиса")
                    _queryLabel.Text = "Тезис:";
                else
                    _queryLabel.Text = "Тема / вопрос:";
            }
        }

        private string GetSelectedTemplateInstruction()
        {
            if (_ragTemplateComboBox == null || _ragTemplateComboBox.SelectedItem == null)
                return string.Empty;

            var template = _ragTemplateComboBox.SelectedItem as RagPromptTemplate;
            return template == null ? string.Empty : template.Instruction;
        }

        private async void GenerateButton_Click(object sender, EventArgs e)
        {
            try
            {
                string textBoxContent = this.PromptTextBox.Text.Trim();
                if (textBoxContent.Length == 0)
                    throw new EmptyTextBoxException(
                        CultureHelper.GetLocalizedString("[GenerateButton_Click] TextBoxEmptyException #1")
                    );

                var rangeBeforeChat = Globals.ThisAddIn.Application.Selection.Range;
                var docRange = Globals.ThisAddIn.Application.ActiveDocument.Range();
                string localCursorContext = GetLocalCursorContext(rangeBeforeChat);
                string templateInstruction = GetSelectedTemplateInstruction();

                var userMessages = new List<UserChatMessage>();

                if (!string.IsNullOrWhiteSpace(localCursorContext))
                {
                    userMessages.Add(
                        new UserChatMessage(
                            "Local Cursor Context / Локальный контекст вокруг курсора. " +
                            "This text comes directly from the current Word document immediately around the insertion point. " +
                            "Prioritize it for transitions, continuation, local coherence, and references such as " +
                            "\"previous paragraph\" or \"next subsection\":\n\n\"" +
                            localCursorContext +
                            "\""
                        )
                    );
                }

                // Put template instructions before the real topic/query. ProcessInformation()
                // intentionally uses messages.Last() as the semantic RAG query, so the
                // retrieval search remains focused on the user's topic instead of generic
                // template wording such as "literature review" or "analytical brief".
                if (!string.IsNullOrWhiteSpace(templateInstruction))
                    userMessages.Add(new UserChatMessage(templateInstruction));

                userMessages.Add(new UserChatMessage(textBoxContent));

                if (rangeBeforeChat.End - rangeBeforeChat.Start > 0)
                    rangeBeforeChat.Delete();

                if (ModelProperties.IsImageModel(ThisAddIn.Model))
                {
                    var streamingAnswer = RAGControl.AskQuestionForImage(
                        new SystemChatMessage(
                            ThisAddIn.SystemPromptLocalization["(GenerateUserControl.cs) _systemPrompt"]
                        ),
                        userMessages,
                        docRange
                    );
                    await Forge.AddStreamingImageContentToRange(streamingAnswer, rangeBeforeChat);
                }
                else
                {
                    var streamingAnswer = RAGControl.AskQuestion(
                        new SystemChatMessage(
                            ThisAddIn.SystemPromptLocalization["(GenerateUserControl.cs) _systemPrompt"]
                        ),
                        userMessages,
                        docRange,
                        GetTemperature()
                    );
                    await Forge.AddStreamingChatContentToRange(streamingAnswer, rangeBeforeChat);
                }

                Globals.ThisAddIn.Application.Selection.SetRange(
                    rangeBeforeChat.Start,
                    rangeBeforeChat.End
                );
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
                    this.GenerateButton.PerformClick();
                }
                else if (e.Control && e.KeyCode == Keys.Back)
                {
                    e.SuppressKeyPress = true;
                    int cursorPosition = this.PromptTextBox.SelectionStart;
                    string text = this.PromptTextBox.Text;

                    while (cursorPosition > 0 && text[cursorPosition - 1] == ' ')
                        cursorPosition--;

                    text = text.TrimEnd();

                    if (string.IsNullOrWhiteSpace(text))
                    {
                        this.PromptTextBox.Clear();
                        this.PromptTextBox.SelectionStart = 0;
                    }
                    else
                    {
                        int lastSpaceIndex = text.LastIndexOf(' ', cursorPosition - 1);
                        if (lastSpaceIndex != -1)
                        {
                            this.PromptTextBox.Text = text.Remove(
                                lastSpaceIndex + 1,
                                cursorPosition - lastSpaceIndex - 1
                            );
                            this.PromptTextBox.SelectionStart = lastSpaceIndex + 1;
                        }
                        else
                        {
                            this.PromptTextBox.Text = text.Remove(0, cursorPosition);
                            this.PromptTextBox.SelectionStart = 0;
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
            this.TemperatureValueLabel.Text = GetTemperature().ToString(
                "0.0",
                Thread.CurrentThread.CurrentUICulture
            );
        }

        private float GetTemperature()
        {
            return this.TemperatureTrackBar.Value / 10f;
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
    }

    public class EmptyTextBoxException : ArgumentException
    {
        public EmptyTextBoxException(string message) : base(message) { }
    }
}
