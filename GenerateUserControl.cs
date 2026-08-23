using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;
using OpenAI.Chat;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    public partial class GenerateUserControl : UserControl
    {
        public static readonly CultureLocalizationHelper CultureHelper = new CultureLocalizationHelper("TextForge.GenerateUserControl", typeof(GenerateUserControl).Assembly);

        private const int DefaultLocalContextBefore = 2;
        private const int DefaultLocalContextAfter = 2;
        private const int MaxLocalContextParagraphs = 20;

        public GenerateUserControl()
        {
            try
            {
                InitializeComponent();
                MatchScrollBarTemperature(); // for floating-point localization
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private async void GenerateButton_Click(object sender, EventArgs e)
        {
            try
            {
                string textBoxContent = this.PromptTextBox.Text;
                if (textBoxContent.Length == 0)
                    throw new EmptyTextBoxException(CultureHelper.GetLocalizedString("[GenerateButton_Click] TextBoxEmptyException #1"));

                /*
                 * Capture the insertion range before starting the request. This also lets
                 * us attach the paragraphs immediately around the cursor. For long Word
                 * documents this local context is more reliable for transitions and
                 * continuation than document-wide semantic retrieval alone.
                 */
                var rangeBeforeChat = Globals.ThisAddIn.Application.Selection.Range;
                var docRange = Globals.ThisAddIn.Application.ActiveDocument.Range();
                string localCursorContext = GetLocalCursorContext(rangeBeforeChat);

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

                // Keep the real user request last. RAGControl.ProcessInformation() uses
                // messages.Last() as the semantic-search query for the whole document/RAG.
                userMessages.Add(new UserChatMessage(textBoxContent));

                // Clear any selected text by the user
                if (rangeBeforeChat.End - rangeBeforeChat.Start > 0)
                    rangeBeforeChat.Delete();

                if (ModelProperties.IsImageModel(ThisAddIn.Model))
                {
                    var streamingAnswer = RAGControl.AskQuestionForImage(
                        new SystemChatMessage(ThisAddIn.SystemPromptLocalization["(GenerateUserControl.cs) _systemPrompt"]),
                        userMessages,
                        docRange
                    );
                    await Forge.AddStreamingImageContentToRange(streamingAnswer, rangeBeforeChat);
                }
                else
                {
                    var streamingAnswer = RAGControl.AskQuestion(
                        new SystemChatMessage(ThisAddIn.SystemPromptLocalization["(GenerateUserControl.cs) _systemPrompt"]),
                        userMessages,
                        docRange,
                        GetTemperature()
                    );
                    await Forge.AddStreamingChatContentToRange(streamingAnswer, rangeBeforeChat);
                }

                Globals.ThisAddIn.Application.Selection.SetRange(rangeBeforeChat.Start, rangeBeforeChat.End);
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

            // Anchor local context to the start of the current selection/cursor.
            localRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            // Include the whole current paragraph, then neighboring paragraphs.
            localRange.Expand(Word.WdUnits.wdParagraph);
            if (paragraphsBefore > 0)
                localRange.MoveStart(Word.WdUnits.wdParagraph, -paragraphsBefore);
            if (paragraphsAfter > 0)
                localRange.MoveEnd(Word.WdUnits.wdParagraph, paragraphsAfter);

            // Prevent an unusually large paragraph from consuming too much model context.
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

                    // Handle multiple trailing spaces
                    while (cursorPosition > 0 && text[cursorPosition - 1] == ' ')
                    {
                        cursorPosition--;
                    }

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
                            // Retain a space after deletion
                            this.PromptTextBox.Text = text.Remove(lastSpaceIndex + 1, cursorPosition - lastSpaceIndex - 1);
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
            } catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void MatchScrollBarTemperature()
        {
            this.TemperatureValueLabel.Text = GetTemperature().ToString("0.0", Thread.CurrentThread.CurrentUICulture);
        }

        private float GetTemperature()
        {
            return this.TemperatureTrackBar.Value / 10f;
        }
    }

    public class EmptyTextBoxException : ArgumentException
    {
        public EmptyTextBoxException(string message) : base(message) { }
    }
}
