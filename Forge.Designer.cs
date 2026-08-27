using System;
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
            // Deliberately small Russian-first UI for the personal scientific-writing build.
            this.ForgeTab.Label = "TextCraft";
            this.ToolsGroup.Label = "Работа с текстом";

            this.ImproveButton.Label = "Улучшить";
            this.ImproveButton.SuperTip = "Сделать выделенный текст яснее и лучше связанным, сохранив смысл и факты.";

            this.FixButton.Label = "Исправить";
            this.FixButton.SuperTip = "Исправить неудачные и неясные формулировки с минимальными изменениями.";

            this.ShortenButton.Label = "Сократить";
            this.ShortenButton.SuperTip = "Сократить выделенный текст примерно на 20%, сохранив факты, термины и ссылки.";

            this.ContinueButton.Label = "Продолжить";
            this.ContinueButton.SuperTip = "Продолжить текст в месте курсора, используя ближайшие абзацы и контекст документа.";

            this.GenerateButton.Label = "Спросить по документу";
            this.GenerateButton.SuperTip = "Открыть простое поле запроса к текущему документу. Ответ вставляется в место курсора.";

            this.GrammarButton.Label = "Грамматика и орфография";
            this.GrammarButton.SuperTip = "Исправить только орфографию, грамматику, пунктуацию и опечатки в выделенном тексте.";

            this.SettingsGroup.Label = "Настройки";
            this.RAGControlButton.Label = "PDF / RAG";
            this.RAGControlButton.SuperTip = "Добавить или удалить PDF-файлы для локального поиска по источникам.";
            this.ModelListDropDown.Label = "Модель";
            this.ModelListDropDown.SuperTip = "Выбрать локальную языковую модель TextCraft.";
            this.DefaultCheckBox.Label = "По умолчанию";
            this.DefaultCheckBox.SuperTip = "Использовать выбранную модель по умолчанию.";

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
            this.ContinueButton = this.Factory.CreateRibbonButton();
            this.GenerateButton = this.Factory.CreateRibbonButton();
            this.GrammarButton = this.Factory.CreateRibbonButton();
            this.SettingsGroup = this.Factory.CreateRibbonGroup();
            this.RAGControlButton = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.ModelListDropDown = this.Factory.CreateRibbonDropDown();
            this.DefaultCheckBox = this.Factory.CreateRibbonCheckBox();
            this.InfoGroup = this.Factory.CreateRibbonGroup();
            this.AboutButton = this.Factory.CreateRibbonButton();
            this.OptionsGroup = this.Factory.CreateRibbonGroup();
            this.CancelButton = this.Factory.CreateRibbonButton();
            this.ForgeTab.SuspendLayout();
            this.ToolsGroup.SuspendLayout();
            this.SettingsGroup.SuspendLayout();
            this.InfoGroup.SuspendLayout();
            this.OptionsGroup.SuspendLayout();
            this.SuspendLayout();

            // ForgeTab
            this.ForgeTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.ForgeTab.Groups.Add(this.ToolsGroup);
            this.ForgeTab.Groups.Add(this.SettingsGroup);
            this.ForgeTab.Groups.Add(this.InfoGroup);
            this.ForgeTab.Groups.Add(this.OptionsGroup);
            this.ForgeTab.Label = "TextCraft";
            this.ForgeTab.Name = "ForgeTab";

            // ToolsGroup
            this.ToolsGroup.Items.Add(this.ImproveButton);
            this.ToolsGroup.Items.Add(this.FixButton);
            this.ToolsGroup.Items.Add(this.ShortenButton);
            this.ToolsGroup.Items.Add(this.ContinueButton);
            this.ToolsGroup.Items.Add(this.GenerateButton);
            this.ToolsGroup.Items.Add(this.GrammarButton);
            this.ToolsGroup.Label = "Работа с текстом";
            this.ToolsGroup.Name = "ToolsGroup";

            // ImproveButton
            this.ImproveButton.Image = global::TextForge.Properties.Resources.counterclockwise_arrows_button_high_contrast;
            this.ImproveButton.Label = "Улучшить";
            this.ImproveButton.Name = "ImproveButton";
            this.ImproveButton.ShowImage = true;
            this.ImproveButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ImproveButton_Click);

            // FixButton
            this.FixButton.Image = global::TextForge.Properties.Resources.memo_high_contrast;
            this.FixButton.Label = "Исправить";
            this.FixButton.Name = "FixButton";
            this.FixButton.ShowImage = true;
            this.FixButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FixButton_Click);

            // ShortenButton
            this.ShortenButton.Image = global::TextForge.Properties.Resources.clipboard_high_contrast;
            this.ShortenButton.Label = "Сократить";
            this.ShortenButton.Name = "ShortenButton";
            this.ShortenButton.ShowImage = true;
            this.ShortenButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShortenButton_Click);

            // ContinueButton
            this.ContinueButton.Image = global::TextForge.Properties.Resources.pen_high_contrast;
            this.ContinueButton.Label = "Продолжить";
            this.ContinueButton.Name = "ContinueButton";
            this.ContinueButton.ShowImage = true;
            this.ContinueButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ContinueButton_Click);

            // GenerateButton / Ask document
            this.GenerateButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.GenerateButton.Image = global::TextForge.Properties.Resources.pen_high_contrast;
            this.GenerateButton.Label = "Спросить по документу";
            this.GenerateButton.Name = "GenerateButton";
            this.GenerateButton.ShowImage = true;
            this.GenerateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GenerateButton_Click);

            // GrammarButton
            this.GrammarButton.Image = global::TextForge.Properties.Resources.face_with_monocle_high_contrast;
            this.GrammarButton.Label = "Грамматика и орфография";
            this.GrammarButton.Name = "GrammarButton";
            this.GrammarButton.ShowImage = true;
            this.GrammarButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GrammarButton_Click);

            // SettingsGroup
            this.SettingsGroup.Items.Add(this.RAGControlButton);
            this.SettingsGroup.Items.Add(this.separator2);
            this.SettingsGroup.Items.Add(this.ModelListDropDown);
            this.SettingsGroup.Items.Add(this.DefaultCheckBox);
            this.SettingsGroup.Label = "Настройки";
            this.SettingsGroup.Name = "SettingsGroup";

            // RAGControlButton
            this.RAGControlButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.RAGControlButton.Image = global::TextForge.Properties.Resources.gear_high_contrast;
            this.RAGControlButton.Label = "PDF / RAG";
            this.RAGControlButton.Name = "RAGControlButton";
            this.RAGControlButton.ShowImage = true;
            this.RAGControlButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RAGControlButton_Click);

            // separator2
            this.separator2.Name = "separator2";

            // ModelListDropDown
            this.ModelListDropDown.Label = "Модель";
            this.ModelListDropDown.Name = "ModelListDropDown";
            this.ModelListDropDown.ShowLabel = false;
            this.ModelListDropDown.SizeString = "XXXXXXXXXXXXXXXXXXXXXXXXX";
            this.ModelListDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ModelListDropDown_SelectionChanged);

            // DefaultCheckBox
            this.DefaultCheckBox.Label = "По умолчанию";
            this.DefaultCheckBox.Name = "DefaultCheckBox";
            this.DefaultCheckBox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DefaultCheckBox_Click);

            // InfoGroup
            this.InfoGroup.Items.Add(this.AboutButton);
            this.InfoGroup.Label = "Информация";
            this.InfoGroup.Name = "InfoGroup";

            // AboutButton
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

            // CancelButton
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
            this.SettingsGroup.ResumeLayout(false);
            this.SettingsGroup.PerformLayout();
            this.InfoGroup.ResumeLayout(false);
            this.InfoGroup.PerformLayout();
            this.OptionsGroup.ResumeLayout(false);
            this.OptionsGroup.PerformLayout();
            this.ResumeLayout(false);
        }

        #endregion

        // --- Simple direct actions -------------------------------------------------

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

        private async void GrammarButton_Click(object sender, RibbonControlEventArgs e)
        {
            await RunQuickSelectionAction(
                "Исправь только орфографию, грамматику, пунктуацию, опечатки и явные ошибки согласования. " +
                "Не переписывай правильные предложения и не меняй стиль без необходимости. " +
                "Сохрани смысл, факты, числа, имена, термины и ссылки. Верни только исправленный текст.",
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ContinueButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GenerateButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GrammarButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup SettingsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RAGControlButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
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
