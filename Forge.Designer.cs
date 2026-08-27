using System;
using System.ClientModel;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using OpenAI.Chat;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    partial class Forge : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private System.ComponentModel.IContainer components = null;

        private const string QuickTextSystemPrompt =
            "Ты помощник по работе с текстом в Microsoft Word. " +
            "Выполняй задачу точно и без лишних пояснений. " +
            "Сохраняй смысл, факты, числа, имена, термины и ссылки. " +
            "Не выдумывай отсутствующие сведения. " +
            "Учитывай стиль исходного текста. " +
            "Если требуется готовый текст, верни только готовый текст.";

        public Forge()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
            Localize();
        }

        private void Localize()
        {
            this.ForgeTab.Label = "TextCraft";
            this.ToolsGroup.Label = "Работа с текстом";
            this.SourcesGroup.Label = "Источники";
            this.SettingsGroup.Label = "Модель";

            this.ImproveButton.Label = "Улучшить";
            this.ImproveButton.SuperTip = "Сделать выделенный текст яснее и лучше связанным, сохранив смысл и факты.";

            this.FixButton.Label = "Исправить";
            this.FixButton.SuperTip = "Исправить неудачные и неясные формулировки с минимальными изменениями.";

            this.ShortenButton.Label = "Сократить";
            this.ShortenButton.SuperTip = "Сократить выделенный текст примерно на 20%, сохранив факты, термины и ссылки.";

            this.ExpandButton.Label = "Расширить";
            this.ExpandButton.SuperTip = "Сделать выделенный текст длиннее за счёт раскрытия уже содержащихся мыслей, без добавления новых фактов.";

            this.KeywordsButton.Label = "Ключевые слова";
            this.KeywordsButton.SuperTip = "Выделить ключевые термины и словосочетания. Исходный текст сохраняется, список добавляется после него.";

            this.GrammarButton.Label = "Грамматика";
            this.GrammarButton.SuperTip = "Исправить орфографию, грамматику, пунктуацию и опечатки без стилистического переписывания.";

            this.ScientificStyleButton.Label = "Научный стиль";
            this.ScientificStyleButton.SuperTip = "Сделать выделенный текст более строгим и академичным без добавления новых фактов.";

            this.ContinueButton.Label = "Продолжить";
            this.ContinueButton.SuperTip = "Продолжить текст в месте курсора, используя ближайшие абзацы и контекст документа.";

            this.GenerateButton.Label = "Спросить";
            this.GenerateButton.SuperTip = "Открыть простое поле запроса к текущему документу и подключенной литературе. Ответ вставляется в место курсора.";

            this.TranslateMenu.Label = "Перевод";
            this.TranslateMenu.SuperTip = "Перевести выделенный текст на выбранный язык с сохранением научной терминологии, чисел и ссылок.";

            this.RAGControlButton.Label = "Литература";
            this.RAGControlButton.SuperTip = "Добавить или удалить PDF-файлы для локального поиска по литературе и источникам.";

            this.StatusLabel.Label = "● Готово";
            this.ModelListDropDown.Label = "Модель";
            this.ModelListDropDown.SuperTip = "Выбрать локальную языковую модель TextCraft. Выбор автоматически сохраняется.";

            this.DefaultCheckBox.Visible = false;

            this.InfoGroup.Label = "Информация";
            this.AboutButton.Label = "О программе";
            this.OptionsGroup.Label = "Выполнение";
            this.CancelButton.Label = "Стоп";
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        private void InitializeComponent()
        {
            this.ForgeTab = this.Factory.CreateRibbonTab();
            this.ToolsGroup = this.Factory.CreateRibbonGroup();
            this.ImproveButton = this.Factory.CreateRibbonButton();
            this.FixButton = this.Factory.CreateRibbonButton();
            this.ShortenButton = this.Factory.CreateRibbonButton();
            this.ExpandButton = this.Factory.CreateRibbonButton();
            this.KeywordsButton = this.Factory.CreateRibbonButton();
            this.GrammarButton = this.Factory.CreateRibbonButton();
            this.ScientificStyleButton = this.Factory.CreateRibbonButton();
            this.ContinueButton = this.Factory.CreateRibbonButton();
            this.GenerateButton = this.Factory.CreateRibbonButton();
            this.TranslateMenu = this.Factory.CreateRibbonMenu();
            this.TranslateRussianButton = this.Factory.CreateRibbonButton();
            this.TranslateEnglishButton = this.Factory.CreateRibbonButton();
            this.TranslateGermanButton = this.Factory.CreateRibbonButton();
            this.TranslateFrenchButton = this.Factory.CreateRibbonButton();
            this.TranslateSpanishButton = this.Factory.CreateRibbonButton();
            this.TranslateItalianButton = this.Factory.CreateRibbonButton();
            this.TranslatePortugueseButton = this.Factory.CreateRibbonButton();
            this.TranslateChineseButton = this.Factory.CreateRibbonButton();
            this.TranslateJapaneseButton = this.Factory.CreateRibbonButton();
            this.TranslateUkrainianButton = this.Factory.CreateRibbonButton();
            this.SourcesGroup = this.Factory.CreateRibbonGroup();
            this.RAGControlButton = this.Factory.CreateRibbonButton();
            this.SettingsGroup = this.Factory.CreateRibbonGroup();
            this.StatusLabel = this.Factory.CreateRibbonLabel();
            this.ModelListDropDown = this.Factory.CreateRibbonDropDown();
            this.DefaultCheckBox = this.Factory.CreateRibbonCheckBox();
            this.InfoGroup = this.Factory.CreateRibbonGroup();
            this.AboutButton = this.Factory.CreateRibbonButton();
            this.OptionsGroup = this.Factory.CreateRibbonGroup();
            this.CancelButton = this.Factory.CreateRibbonButton();
            this.ForgeTab.SuspendLayout();
            this.ToolsGroup.SuspendLayout();
            this.SourcesGroup.SuspendLayout();
            this.SettingsGroup.SuspendLayout();
            this.InfoGroup.SuspendLayout();
            this.OptionsGroup.SuspendLayout();
            this.SuspendLayout();

            // ForgeTab
            this.ForgeTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.ForgeTab.Groups.Add(this.ToolsGroup);
            this.ForgeTab.Groups.Add(this.SourcesGroup);
            this.ForgeTab.Groups.Add(this.SettingsGroup);
            this.ForgeTab.Groups.Add(this.InfoGroup);
            this.ForgeTab.Groups.Add(this.OptionsGroup);
            this.ForgeTab.Label = "TextCraft";
            this.ForgeTab.Name = "ForgeTab";

            // ToolsGroup
            this.ToolsGroup.Items.Add(this.ImproveButton);
            this.ToolsGroup.Items.Add(this.FixButton);
            this.ToolsGroup.Items.Add(this.ShortenButton);
            this.ToolsGroup.Items.Add(this.ExpandButton);
            this.ToolsGroup.Items.Add(this.KeywordsButton);
            this.ToolsGroup.Items.Add(this.GrammarButton);
            this.ToolsGroup.Items.Add(this.ScientificStyleButton);
            this.ToolsGroup.Items.Add(this.ContinueButton);
            this.ToolsGroup.Items.Add(this.GenerateButton);
            this.ToolsGroup.Items.Add(this.TranslateMenu);
            this.ToolsGroup.Label = "Работа с текстом";
            this.ToolsGroup.Name = "ToolsGroup";

            // ImproveButton
            this.ImproveButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ImproveButton.Image = global::TextForge.Properties.Resources.counterclockwise_arrows_button_high_contrast;
            this.ImproveButton.Label = "Улучшить";
            this.ImproveButton.Name = "ImproveButton";
            this.ImproveButton.ShowImage = true;
            this.ImproveButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ImproveButton_Click);

            // FixButton
            this.FixButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.FixButton.Image = global::TextForge.Properties.Resources.memo_high_contrast;
            this.FixButton.Label = "Исправить";
            this.FixButton.Name = "FixButton";
            this.FixButton.ShowImage = true;
            this.FixButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FixButton_Click);

            // ShortenButton
            this.ShortenButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ShortenButton.Image = global::TextForge.Properties.Resources.clipboard_high_contrast;
            this.ShortenButton.Label = "Сократить";
            this.ShortenButton.Name = "ShortenButton";
            this.ShortenButton.ShowImage = true;
            this.ShortenButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShortenButton_Click);

            // ExpandButton
            this.ExpandButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ExpandButton.Image = global::TextForge.Properties.Resources.pen_high_contrast;
            this.ExpandButton.Label = "Расширить";
            this.ExpandButton.Name = "ExpandButton";
            this.ExpandButton.ShowImage = true;
            this.ExpandButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExpandButton_Click);

            // KeywordsButton
            this.KeywordsButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.KeywordsButton.Image = global::TextForge.Properties.Resources.face_with_monocle_high_contrast;
            this.KeywordsButton.Label = "Ключевые слова";
            this.KeywordsButton.Name = "KeywordsButton";
            this.KeywordsButton.ShowImage = true;
            this.KeywordsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.KeywordsButton_Click);

            // GrammarButton
            this.GrammarButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.GrammarButton.Image = global::TextForge.Properties.Resources.face_with_monocle_high_contrast;
            this.GrammarButton.Label = "Грамматика";
            this.GrammarButton.Name = "GrammarButton";
            this.GrammarButton.ShowImage = true;
            this.GrammarButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GrammarButton_Click);

            // ScientificStyleButton
            this.ScientificStyleButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ScientificStyleButton.Image = global::TextForge.Properties.Resources.clipboard_high_contrast;
            this.ScientificStyleButton.Label = "Научный стиль";
            this.ScientificStyleButton.Name = "ScientificStyleButton";
            this.ScientificStyleButton.ShowImage = true;
            this.ScientificStyleButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ScientificStyleButton_Click);

            // ContinueButton
            this.ContinueButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ContinueButton.Image = global::TextForge.Properties.Resources.pen_high_contrast;
            this.ContinueButton.Label = "Продолжить";
            this.ContinueButton.Name = "ContinueButton";
            this.ContinueButton.ShowImage = true;
            this.ContinueButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ContinueButton_Click);

            // GenerateButton
            this.GenerateButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.GenerateButton.Image = global::TextForge.Properties.Resources.pen_high_contrast;
            this.GenerateButton.Label = "Спросить";
            this.GenerateButton.Name = "GenerateButton";
            this.GenerateButton.ShowImage = true;
            this.GenerateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GenerateButton_Click);

            // TranslateMenu
            this.TranslateMenu.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.TranslateMenu.Image = global::TextForge.Properties.Resources.counterclockwise_arrows_button_high_contrast;
            this.TranslateMenu.Items.Add(this.TranslateRussianButton);
            this.TranslateMenu.Items.Add(this.TranslateEnglishButton);
            this.TranslateMenu.Items.Add(this.TranslateGermanButton);
            this.TranslateMenu.Items.Add(this.TranslateFrenchButton);
            this.TranslateMenu.Items.Add(this.TranslateSpanishButton);
            this.TranslateMenu.Items.Add(this.TranslateItalianButton);
            this.TranslateMenu.Items.Add(this.TranslatePortugueseButton);
            this.TranslateMenu.Items.Add(this.TranslateChineseButton);
            this.TranslateMenu.Items.Add(this.TranslateJapaneseButton);
            this.TranslateMenu.Items.Add(this.TranslateUkrainianButton);
            this.TranslateMenu.Label = "Перевод";
            this.TranslateMenu.Name = "TranslateMenu";
            this.TranslateMenu.ShowImage = true;

            ConfigureTranslationButton(this.TranslateRussianButton, "TranslateRussianButton", "Русский");
            ConfigureTranslationButton(this.TranslateEnglishButton, "TranslateEnglishButton", "Английский");
            ConfigureTranslationButton(this.TranslateGermanButton, "TranslateGermanButton", "Немецкий");
            ConfigureTranslationButton(this.TranslateFrenchButton, "TranslateFrenchButton", "Французский");
            ConfigureTranslationButton(this.TranslateSpanishButton, "TranslateSpanishButton", "Испанский");
            ConfigureTranslationButton(this.TranslateItalianButton, "TranslateItalianButton", "Итальянский");
            ConfigureTranslationButton(this.TranslatePortugueseButton, "TranslatePortugueseButton", "Португальский");
            ConfigureTranslationButton(this.TranslateChineseButton, "TranslateChineseButton", "Китайский");
            ConfigureTranslationButton(this.TranslateJapaneseButton, "TranslateJapaneseButton", "Японский");
            ConfigureTranslationButton(this.TranslateUkrainianButton, "TranslateUkrainianButton", "Украинский");

            // SourcesGroup
            this.SourcesGroup.Items.Add(this.RAGControlButton);
            this.SourcesGroup.Label = "Источники";
            this.SourcesGroup.Name = "SourcesGroup";

            this.RAGControlButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.RAGControlButton.Image = global::TextForge.Properties.Resources.memo_high_contrast;
            this.RAGControlButton.Label = "Литература";
            this.RAGControlButton.Name = "RAGControlButton";
            this.RAGControlButton.ShowImage = true;
            this.RAGControlButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RAGControlButton_Click);

            // SettingsGroup
            this.SettingsGroup.Items.Add(this.StatusLabel);
            this.SettingsGroup.Items.Add(this.ModelListDropDown);
            this.SettingsGroup.Label = "Модель";
            this.SettingsGroup.Name = "SettingsGroup";

            this.StatusLabel.Label = "● Готово";
            this.StatusLabel.Name = "StatusLabel";

            this.ModelListDropDown.Label = "Модель";
            this.ModelListDropDown.Name = "ModelListDropDown";
            this.ModelListDropDown.ShowLabel = false;
            this.ModelListDropDown.SizeString = "XXXXXXXXXXXXXXXXXXXXXXXXX";
            this.ModelListDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ModelListDropDown_SelectionChanged);

            this.DefaultCheckBox.Label = "По умолчанию";
            this.DefaultCheckBox.Name = "DefaultCheckBox";
            this.DefaultCheckBox.Visible = false;
            this.DefaultCheckBox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DefaultCheckBox_Click);

            // InfoGroup
            this.InfoGroup.Items.Add(this.AboutButton);
            this.InfoGroup.Label = "Информация";
            this.InfoGroup.Name = "InfoGroup";

            this.AboutButton.Image = global::TextForge.Properties.Resources.information_high_contrast;
            this.AboutButton.Label = "О программе";
            this.AboutButton.Name = "AboutButton";
            this.AboutButton.ShowImage = true;
            this.AboutButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AboutButton_Click);

            // OptionsGroup
            this.OptionsGroup.Items.Add(this.CancelButton);
            this.OptionsGroup.Label = "Выполнение";
            this.OptionsGroup.Name = "OptionsGroup";
            this.OptionsGroup.Visible = false;

            this.CancelButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.CancelButton.Image = global::TextForge.Properties.Resources.stop_sign_flat;
            this.CancelButton.Label = "Стоп";
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.ShowImage = true;
            this.CancelButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CancelButton_Click);

            // Forge
            this.Name = "Forge";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.ForgeTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Forge_Load);
            this.ForgeTab.ResumeLayout(false);
            this.ForgeTab.PerformLayout();
            this.ToolsGroup.ResumeLayout(false);
            this.ToolsGroup.PerformLayout();
            this.SourcesGroup.ResumeLayout(false);
            this.SourcesGroup.PerformLayout();
            this.SettingsGroup.ResumeLayout(false);
            this.SettingsGroup.PerformLayout();
            this.InfoGroup.ResumeLayout(false);
            this.InfoGroup.PerformLayout();
            this.OptionsGroup.ResumeLayout(false);
            this.OptionsGroup.PerformLayout();
            this.ResumeLayout(false);
        }

        private void ConfigureTranslationButton(RibbonButton button, string name, string label)
        {
            button.Label = label;
            button.Name = name;
            button.ShowImage = false;
            button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TranslateButton_Click);
        }

        #endregion

        private async void ImproveButton_Click(object sender, RibbonControlEventArgs e)
        {
            await RunQuickSelectionAction(
                "Сделай выделенный текст яснее, естественнее и лучше связанным. " +
                "Сохрани научный стиль, смысл, факты, числа, термины и ссылки. " +
                "Не добавляй новых сведений. Верни только улучшенный текст.",
                0.10f
            );
        }

        private async void FixButton_Click(object sender, RibbonControlEventArgs e)
        {
            await RunQuickSelectionAction(
                "Исправь неудачные, тяжёлые или неясные формулировки в выделенном тексте. " +
                "Делай минимальные изменения. Не меняй смысл, факты, числа, термины и ссылки. " +
                "Верни только исправленный текст.",
                0.08f
            );
        }

        private async void ShortenButton_Click(object sender, RibbonControlEventArgs e)
        {
            await RunQuickSelectionAction(
                "Сократи выделенный текст примерно на 20 процентов. " +
                "Убери повторы и лишние слова, но сохрани смысл, факты, числа, термины и ссылки. " +
                "Не добавляй новых сведений. Верни только сокращённый текст.",
                0.08f
            );
        }

        private async void ExpandButton_Click(object sender, RibbonControlEventArgs e)
        {
            await RunQuickSelectionAction(
                "Расширь выделенный текст примерно на 40–60 процентов. " +
                "Раскрой уже содержащиеся мысли, уточни логические связи, сделай переходы между предложениями более явными и при необходимости поясни уже названные понятия. " +
                "Не добавляй новых фактов, чисел, результатов исследований, причинно-следственных утверждений, ссылок или источников, которых нет в исходном тексте. " +
                "Сохрани научный стиль, терминологию и фактический смысл. Верни только расширенный текст.",
                0.06f
            );
        }

        private async void KeywordsButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                await InsertKeywordsAfterSelectionAsync();
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

        private static async Task InsertKeywordsAfterSelectionAsync()
        {
            Word.Selection selection = Globals.ThisAddIn.Application.Selection;
            if (selection == null || selection.End <= selection.Start)
            {
                MessageBox.Show(
                    "Сначала выделите текст, из которого нужно извлечь ключевые слова.",
                    "TextCraft",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return;
            }

            Word.Range sourceRange = selection.Range.Duplicate;
            string selectedText = sourceRange.Text ?? string.Empty;

            ChatClient client = new ChatClient(
                ThisAddIn.Model,
                new ApiKeyCredential(ThisAddIn.ApiKey),
                ThisAddIn.ClientOptions
            );

            var messages = new List<ChatMessage>()
            {
                new SystemChatMessage(
                    "Ты помощник по научному тексту. Извлекай только ключевые термины и устойчивые словосочетания из предоставленного текста. " +
                    "Не придумывай отсутствующие понятия и не добавляй пояснений."
                ),
                new UserChatMessage(
                    "Извлеки 5–12 ключевых слов или ключевых словосочетаний из текста. " +
                    "Выбирай наиболее содержательные термины, отражающие предмет, объект, методы, процессы и основные результаты, если они присутствуют. " +
                    "Сохрани язык исходного текста. Не включай слишком общие слова. " +
                    "Верни только список через точку с запятой, без нумерации, Markdown и вводных слов.\n\nТекст:\n" + selectedText
                )
            };

            var streamingAnswer = client.CompleteChatStreamingAsync(
                messages,
                new ChatCompletionOptions() { Temperature = 0.10f },
                ThisAddIn.CancellationTokenSource.Token
            );

            Word.Range insertionRange = sourceRange.Duplicate;
            insertionRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            insertionRange.Text = Environment.NewLine + "Ключевые слова: ";
            insertionRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            await AddStreamingChatContentToRange(streamingAnswer, insertionRange);
            Globals.ThisAddIn.Application.Selection.SetRange(insertionRange.Start, insertionRange.End);
        }

        private async void GrammarButton_Click(object sender, RibbonControlEventArgs e)
        {
            await RunQuickSelectionAction(
                "Исправь только орфографию, грамматику, пунктуацию, опечатки и явные ошибки согласования. " +
                "Не переписывай правильные предложения и не меняй стиль без необходимости. " +
                "Сохрани смысл, факты, числа, имена, термины и ссылки. Верни только исправленный текст.",
                0.05f
            );
        }

        private async void ScientificStyleButton_Click(object sender, RibbonControlEventArgs e)
        {
            await RunQuickSelectionAction(
                "Перепиши выделенный текст в строгом научном академическом стиле. " +
                "Убери разговорные, оценочные и расплывчатые формулировки, сделай изложение точным, нейтральным и логически связным. " +
                "Не добавляй новых фактов и не меняй фактический смысл. " +
                "Сохрани числа, формулы, единицы измерения, имена, специальные термины, цитаты и библиографические ссылки. " +
                "Верни только итоговый текст.",
                0.07f
            );
        }

        private async void TranslateButton_Click(object sender, RibbonControlEventArgs e)
        {
            string targetLanguage;
            switch (e.Control.Id)
            {
                case "TranslateRussianButton": targetLanguage = "русский"; break;
                case "TranslateEnglishButton": targetLanguage = "английский"; break;
                case "TranslateGermanButton": targetLanguage = "немецкий"; break;
                case "TranslateFrenchButton": targetLanguage = "французский"; break;
                case "TranslateSpanishButton": targetLanguage = "испанский"; break;
                case "TranslateItalianButton": targetLanguage = "итальянский"; break;
                case "TranslatePortugueseButton": targetLanguage = "португальский"; break;
                case "TranslateChineseButton": targetLanguage = "китайский"; break;
                case "TranslateJapaneseButton": targetLanguage = "японский"; break;
                case "TranslateUkrainianButton": targetLanguage = "украинский"; break;
                default: return;
            }

            await RunQuickSelectionAction(
                "Переведи выделенный текст на " + targetLanguage + " язык. " +
                "Определи исходный язык автоматически. Сохрани смысл и научный регистр. " +
                "Используй общепринятую научную и профессиональную терминологию целевого языка. " +
                "Не меняй числа, формулы, единицы измерения, DOI, URL, библиографические ссылки и обозначения. " +
                "Не добавляй комментариев или пояснений. Верни только перевод.",
                0.05f
            );
        }

        private async Task RunQuickSelectionAction(string instruction, float temperature)
        {
            try
            {
                Word.Selection selection = Globals.ThisAddIn.Application.Selection;
                if (selection == null || selection.End <= selection.Start)
                {
                    MessageBox.Show(
                        "Сначала выделите текст, который нужно обработать.",
                        "TextCraft",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    return;
                }

                await AnalyzeText(QuickTextSystemPrompt, instruction, temperature);
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

        private async void ContinueButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                await ContinueAtCursorAsync();
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

        private static async Task ContinueAtCursorAsync()
        {
            Word.Range insertionRange = Globals.ThisAddIn.Application.Selection.Range.Duplicate;
            insertionRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            Word.Range localRange = insertionRange.Duplicate;
            localRange.Expand(Word.WdUnits.wdParagraph);
            localRange.MoveStart(Word.WdUnits.wdParagraph, -2);
            localRange.MoveEnd(Word.WdUnits.wdParagraph, 2);

            int maxLocalTokens = Math.Max(512, (int)(ThisAddIn.ContextLength * 0.20));
            string localContext = CommonUtils.SubstringTokens(localRange.Text ?? string.Empty, maxLocalTokens);

            string request =
                "Продолжи текущую мысль одним коротким абзацем в стиле документа. " +
                "Опирайся только на факты из текста документа, локального контекста и RAG. " +
                "Не выдумывай новые факты, числа, ссылки или источники. " +
                "Если данных для нового утверждения недостаточно, сделай нейтральный логический переход. " +
                "Верни только текст продолжения.\n\n" +
                "Текст рядом с курсором:\n" + localContext;

            var messages = new List<ChatMessage>()
            {
                new UserChatMessage(request)
            };

            var answer = RAGControl.AskQuestion(
                new SystemChatMessage(QuickTextSystemPrompt),
                messages,
                Globals.ThisAddIn.Application.ActiveDocument.Range(),
                0.08f
            );

            await AddStreamingChatContentToRange(answer, insertionRange);
            Globals.ThisAddIn.Application.Selection.SetRange(insertionRange.Start, insertionRange.End);
        }

        internal Microsoft.Office.Tools.Ribbon.RibbonTab ForgeTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ToolsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ImproveButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton FixButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ShortenButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExpandButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton KeywordsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GrammarButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ScientificStyleButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ContinueButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GenerateButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu TranslateMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TranslateRussianButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TranslateEnglishButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TranslateGermanButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TranslateFrenchButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TranslateSpanishButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TranslateItalianButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TranslatePortugueseButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TranslateChineseButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TranslateJapaneseButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TranslateUkrainianButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup SourcesGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RAGControlButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup SettingsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel StatusLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ModelListDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox DefaultCheckBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup InfoGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AboutButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup OptionsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CancelButton;
    }

    partial class ThisRibbonCollection
    {
        internal Forge Forge
        {
            get { return this.GetRibbon<Forge>(); }
        }
    }
}
