using System;
using System.ClientModel;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;
using OpenAI.Chat;
using OpenAI.Images;
using Task = System.Threading.Tasks.Task;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    public partial class Forge
    {
        // Public
        public static SystemChatMessage CommentSystemPrompt;
        public static readonly CultureLocalizationHelper CultureHelper = new CultureLocalizationHelper("TextForge.Forge", typeof(Forge).Assembly);
        public static readonly object InitializeDoor = new object();

        // Private
        private AboutBox _box;
        private static RibbonGroup _optionsBox;
        private static RibbonLabel _statusLabel;

        private void Forge_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                if (Globals.ThisAddIn.Application.Documents.Count > 0)
                    ThisAddIn.AddTaskPanes(Globals.ThisAddIn.Application.ActiveDocument);

                _statusLabel = this.StatusLabel;
                SetStatus("● Готово");

                Thread startup = new Thread(InitializeForge);
                startup.SetApartmentState(ApartmentState.STA);
                startup.Start();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void InitializeForge()
        {
            try
            {
                lock (InitializeDoor)
                {
                    if (!ThisAddIn.IsAddinInitialized)
                        ThisAddIn.InitializeAddIn();
                    
                    CommentSystemPrompt = new SystemChatMessage(ThisAddIn.SystemPromptLocalization["this.CommentSystemPrompt"]);

                    PopulateDropdownList(ThisAddIn.LanguageModelList);
                }
                _box = new AboutBox();
                _optionsBox = this.OptionsGroup;
                SetStatus("● Готово");
            }
            catch (Exception ex)
            {
                SetStatus("● Ошибка");
                CommonUtils.DisplayError(ex);
            }
        }

        private void PopulateDropdownList(IEnumerable<string> modelList)
        {
            var ribbonFactory = Globals.Factory.GetRibbonFactory();
            var sortedModels = modelList.OrderBy(m => m).ToList();
            foreach (string model in sortedModels)
            {
                {
                    var newItem = ribbonFactory.CreateRibbonDropDownItem();
                    newItem.Label = model;
                    ModelListDropDown.Items.Add(newItem);

                    if (model == ThisAddIn.Model)
                    {
                        ModelListDropDown.SelectedItem = newItem;
                        UpdateCheckbox();
                    }
                }
            }
        }

        private async void ModelListDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string selectedModel = GetSelectedItemLabel();
                ThisAddIn.Model = selectedModel;

                // In the simplified UI the selected model is automatically remembered.
                Properties.Settings.Default.DefaultModel = selectedModel;
                Properties.Settings.Default.Save();
                UpdateCheckbox();

                SetStatus("◌ Модель…");
                await Task.Run(() =>
                {
                    ThisAddIn.ContextLength = ModelProperties.GetContextLength(ThisAddIn.Model, ThisAddIn.ModelList);
                });
                SetStatus("● Готово");
            }
            catch (Exception ex)
            {
                SetStatus("● Ошибка");
                CommonUtils.DisplayError(ex);
            }
        }

        private void DefaultCheckBox_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (this.DefaultCheckBox.Checked)
                    Properties.Settings.Default.DefaultModel = GetSelectedItemLabel();
                else
                    Properties.Settings.Default.DefaultModel = null;
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private string GetSelectedItemLabel()
        {
            return ModelListDropDown.SelectedItem.Label;
        }

        private void GenerateButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var taskpanes = ThisAddIn.AllTaskPanes[Globals.ThisAddIn.Application.ActiveDocument];
                taskpanes.Item1.Visible = !taskpanes.Item1.Visible;
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void RAGControlButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var taskpanes = ThisAddIn.AllTaskPanes[Globals.ThisAddIn.Application.ActiveDocument];
                taskpanes.Item2.Visible = !taskpanes.Item2.Visible;
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void AboutButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                _box.ShowDialog();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }
        private void CancelButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                CancelButtonVisibility(false);
                ThisAddIn.CancellationTokenSource.Cancel();
                ThisAddIn.CancellationTokenSource = new CancellationTokenSource();
                SetStatus("● Готово");
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private async void WritingToolsGallery_ButtonClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                switch (e.Control.Id)
                {
                    case "ReviewButton":
                        await ReviewButton_Click();
                        break;
                    case "ProofreadButton":
                        await ProofreadButton_Click();
                        break;
                    case "RewriteButton":
                        await RewriteButton_Click();
                        break;
                    default:
                        throw new ArgumentOutOfRangeException(CultureHelper.GetLocalizedString("[WritingToolsGallery_ButtonClick] ArgumentOutOfRangeException #1"));
                }
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private static async Task ReviewButton_Click()
        {
            string userPrompt = CultureHelper.GetLocalizedString("[ReviewButton_Click] UserPrompt");
            Word.Paragraphs paragraphs = CommonUtils.GetActiveDocument().Paragraphs;

            bool hasCommented = false;
            if (Globals.ThisAddIn.Application.Selection.End - Globals.ThisAddIn.Application.Selection.Start > 0)
            {
                var selectionRange = CommonUtils.GetSelectionRange();
                try
                {
                    await CommentHandler.AddComment(CommonUtils.GetComments(), selectionRange, Review(paragraphs, selectionRange, userPrompt));
                }
                catch (OperationCanceledException ex)
                {
                    CommonUtils.DisplayWarning(ex);
                }
                hasCommented = true;
            }
            else
            {
                Word.Document document = CommonUtils.GetActiveDocument(); // Hash code of the active document gets changed after each comment!
                foreach (Word.Paragraph p in paragraphs)
                    // It isn't a paragraph if it doesn't contain a full stop.
                    if (ContainsFullStop(p.Range.Text))
                    {
                        await CommentHandler.AddComment(CommonUtils.GetComments(), p.Range, Review(paragraphs, p.Range, userPrompt, document));
                        hasCommented = true;
                    }
            }
            if (!hasCommented)
                MessageBox.Show(CultureHelper.GetLocalizedString("[ReviewButton_Click] MessageBox #1 (text)"), CultureHelper.GetLocalizedString("[ReviewButton_Click] MessageBox #1 (caption)"), MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static bool ContainsFullStop(string value)
        {
            string currentLanguage = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName;
            var languageSpecificFullStops = new Dictionary<string, char[]>
            {
                { "hi", new char[] { '।' } },
                { "am", new char[] { '።' } },
                { "hy", new char[] { '։' } },
                { "zh", new char[] { '。' } },
                { "ja", new char[] { '。' } },
                { "ko", new char[] { '。' } },
                { "ar", new char[] { '۔' } },
            };

            if (value.Contains('.'))
                return true;

            if (languageSpecificFullStops.TryGetValue(currentLanguage, out char[] specificStops))
                return value.IndexOfAny(specificStops) >= 0;

            return false;
        }

        private static async Task ProofreadButton_Click()
        {
            await AnalyzeText(
                ThisAddIn.SystemPromptLocalization["[ProofreadButton_Click] SystemPrompt"],
                CultureHelper.GetLocalizedString("[ProofreadButton_Click] UserPrompt"),
                0.1f
            );
        }

        private static async Task RewriteButton_Click()
        {
            await AnalyzeText(
                ThisAddIn.SystemPromptLocalization["[RewriteButton_Click] SystemPrompt"],
                CultureHelper.GetLocalizedString("[RewriteButton_Click] UserPrompt"),
                0.4f
            );
        }

        private static async Task AnalyzeText(string systemPrompt, string userPrompt, float temperature)
        {
            var selectionRange = Globals.ThisAddIn.Application.Selection.Range;
            var range = (selectionRange.End - selectionRange.Start > 0) ? selectionRange : throw new InvalidRangeException(CultureHelper.GetLocalizedString("[AnalyzeText] InvalidRangeException #1"));
            string selectedText = range.Text;

            ChatClient client = new ChatClient(ThisAddIn.Model, new ApiKeyCredential(ThisAddIn.ApiKey), ThisAddIn.ClientOptions);
            var streamingAnswer = client.CompleteChatStreamingAsync(
                new List<ChatMessage>() { new SystemChatMessage(systemPrompt), new UserChatMessage($@"{CultureHelper.GetLocalizedString("[AnalyzeText] UserChatMessage #1")}:\n{GetTextFromParagraphs(selectionRange.Paragraphs)}"), new UserChatMessage(@$"{userPrompt}:\n{selectedText}") },
                new ChatCompletionOptions() { Temperature = temperature * 2 },
                ThisAddIn.CancellationTokenSource.Token
            );

            range.Delete();
            try
            {
                await AddStreamingChatContentToRange(streamingAnswer, range);
                if (selectedText.EndsWith("\r"))
                    range.Text += Environment.NewLine;
            }
            catch (OperationCanceledException ex)
            {
                CommonUtils.DisplayWarning(ex);
            }
            Globals.ThisAddIn.Application.Selection.SetRange(range.Start, range.End);
        }

        private static string GetTextFromParagraphs(Paragraphs paragraphs)
        {
            StringBuilder textBuilder = new StringBuilder(paragraphs.Count);
            foreach (Paragraph p in paragraphs)
                textBuilder.AppendLine(p.Range.Text);
            return textBuilder.ToString();
        }

        public static async Task AddStreamingChatContentToRange(AsyncCollectionResult<StreamingChatCompletionUpdate> streamingAnswer, Word.Range range)
        {
            StringBuilder response = new StringBuilder();
            CancelButtonVisibility(true);
            try
            {
                await foreach (var update in streamingAnswer.WithCancellation(ThisAddIn.CancellationTokenSource.Token))
                {
                    if (ThisAddIn.CancellationTokenSource.IsCancellationRequested) break;
                    foreach (var newContent in update.ContentUpdate)
                    {
                        switch (newContent.Kind)
                        {
                            case ChatMessageContentPartKind.Text:
                                range.Text += newContent.Text;
                                response.Append(newContent.Text);
                                break;
                            case ChatMessageContentPartKind.Refusal:
                                MessageBox.Show(CultureHelper.GetLocalizedString("[AddStreamingChatContentToRange] MessageBox Text #1"), CultureHelper.GetLocalizedString("[AddStreamingChatContentToRange] MessageBox Caption #1"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                break;
                            default:
                                break;
                        }
                    }
                }
            }
            finally
            {
                CancelButtonVisibility(false);
            }

            range.Text = WordMarkdown.RemoveMarkdownSyntax(response.ToString());
            WordMarkdown.ApplyAllMarkdownFormatting(range, response.ToString());
        }

        public static async Task AddStreamingImageContentToRange(Task<ClientResult<GeneratedImage>> streamingAnswer, Word.Range range)
        {
            StringBuilder response = new StringBuilder();
            CancelButtonVisibility(true);
            try
            {
                try
                {
                    ClientResult<GeneratedImage> clientResult = await streamingAnswer;
                    string pictureAddress = GetPictureAddress(clientResult);
                    range.InlineShapes.AddPicture(pictureAddress);
                    File.Delete(pictureAddress);
                }
                catch (Exception ex)
                {
                    CommonUtils.DisplayError(CultureHelper.GetLocalizedString("[AddStreamingImageContentToRange] Exception #1"), ex);
                }
            }
            finally
            {
                CancelButtonVisibility(false);
            }

            range.Text = WordMarkdown.RemoveMarkdownSyntax(response.ToString());
            WordMarkdown.ApplyAllMarkdownFormatting(range, response.ToString());
        }

        public static void CancelButtonVisibility(bool option)
        {
            if (_optionsBox != null)
                _optionsBox.Visible = option;
            SetStatus(option ? "◌ Думает…" : "● Готово");
        }

        private static void SetStatus(string value)
        {
            try
            {
                if (_statusLabel != null)
                    _statusLabel.Label = value;
            }
            catch
            {
                // Status is informational only and must never interrupt editing.
            }
        }

        private void UpdateCheckbox()
        {
            DefaultCheckBox.Checked = (Properties.Settings.Default.DefaultModel == ThisAddIn.Model);
        }

        private static AsyncCollectionResult<StreamingChatCompletionUpdate> Review(Word.Paragraphs context, Word.Range p, string userPrompt, Word.Document doc = null)
        {
            var docRange = Globals.ThisAddIn.Application.ActiveDocument.Range();
            List<UserChatMessage> chatHistory = new List<UserChatMessage>()
            {
                new UserChatMessage($@"{CultureHelper.GetLocalizedString("[Review] chatHistory #1")}\n\"{CommonUtils.SubstringTokens(p.Text, (int)(ThisAddIn.ContextLength * 0.2))}\""),
                new UserChatMessage(userPrompt)
            };
            return RAGControl.AskQuestion(CommentSystemPrompt, chatHistory, docRange, 0.5f, doc);
        }

        public static string GetPictureAddress(GeneratedImage newContent)
        {
            if (newContent.ImageBytes != null)
            {
                string tempFilePath = Path.GetTempFileName();
                File.WriteAllBytes(tempFilePath, newContent.ImageBytes.ToArray());
                return tempFilePath;
            }
            else if (!string.IsNullOrEmpty(newContent.ImageUri.ToString()))
            {
                throw new InvalidDataException(CultureHelper.GetLocalizedString("[GetPictureAddress] InvalidDataException #1"));
            }
            else
            {
                throw new InvalidOperationException(CultureHelper.GetLocalizedString("[GetPictureAddress] InvalidOperationException #1"));
            }
        }
    }
}
