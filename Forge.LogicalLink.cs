using System;
using System.ClientModel;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using OpenAI.Chat;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    public partial class Forge
    {
        private Word.Range _logicalLinkFirstRange;
        private string _logicalLinkFirstText;
        private string _logicalLinkDocumentIdentity;

        private async void LogicalLinkButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
                Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
                string currentText = NormalizeLogicalLinkText(currentRange.Text);

                if (currentRange.End <= currentRange.Start || string.IsNullOrWhiteSpace(currentText))
                {
                    MessageBox.Show(
                        _logicalLinkFirstRange == null
                            ? "Выделите первый фрагмент текста и нажмите «Связать». Затем выделите второй фрагмент и нажмите кнопку ещё раз."
                            : "Теперь выделите второй фрагмент текста и снова нажмите «Связать».",
                        "НеZнайка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    return;
                }

                string documentIdentity = GetLogicalLinkDocumentIdentity(document);

                if (_logicalLinkFirstRange == null)
                {
                    _logicalLinkFirstRange = currentRange.Duplicate;
                    _logicalLinkFirstText = currentText;
                    _logicalLinkDocumentIdentity = documentIdentity;
                    LogicalLinkButton.Label = "2-я часть";
                    SetStatus("● Выберите 2-ю часть");
                    return;
                }

                if (!string.Equals(
                        _logicalLinkDocumentIdentity,
                        documentIdentity,
                        StringComparison.OrdinalIgnoreCase
                    ))
                {
                    ResetLogicalLinkSelection();
                    MessageBox.Show(
                        "Фрагменты должны находиться в одном документе. Выделите первый фрагмент заново.",
                        "НеZнайка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    return;
                }

                Word.Range secondRange = currentRange.Duplicate;
                int firstStart = _logicalLinkFirstRange.Start;
                int firstEnd = _logicalLinkFirstRange.End;
                int secondStart = secondRange.Start;
                int secondEnd = secondRange.End;

                if (firstStart < secondEnd && secondStart < firstEnd)
                {
                    MessageBox.Show(
                        "Второй фрагмент пересекается с первым. Выделите другую часть текста.",
                        "НеZнайка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    return;
                }

                string firstText = _logicalLinkFirstText;
                try
                {
                    string liveFirstText = NormalizeLogicalLinkText(_logicalLinkFirstRange.Text);
                    if (!string.IsNullOrWhiteSpace(liveFirstText))
                        firstText = liveFirstText;
                }
                catch
                {
                    // Keep the captured text if Word can no longer read the live Range text.
                }

                string earlierText;
                string laterText;
                Word.Range laterRange;

                if (firstStart <= secondStart)
                {
                    earlierText = firstText;
                    laterText = currentText;
                    laterRange = secondRange;
                }
                else
                {
                    earlierText = currentText;
                    laterText = firstText;
                    laterRange = _logicalLinkFirstRange;
                }

                string transition = await GenerateLogicalTransitionAsync(earlierText, laterText);
                transition = NormalizeGeneratedTransition(transition);
                if (string.IsNullOrWhiteSpace(transition))
                    return;

                int insertPosition = laterRange.Start;
                Word.Range insertRange = document.Range(insertPosition, insertPosition);
                insertRange.Text = transition + " ";
                Globals.ThisAddIn.Application.Selection.SetRange(insertRange.Start, insertRange.End);

                ResetLogicalLinkSelection();
                SetStatus("● Связано");
            }
            catch (OperationCanceledException)
            {
                SetStatus("● Остановлено");
            }
            catch (Exception ex)
            {
                ResetLogicalLinkSelection();
                SetStatus("● Ошибка");
                CommonUtils.DisplayError(ex);
            }
        }

        private static async Task<string> GenerateLogicalTransitionAsync(
            string firstFragment,
            string secondFragment
        )
        {
            ChatClient client = new ChatClient(
                ThisAddIn.Model,
                new ApiKeyCredential(ThisAddIn.ApiKey),
                ThisAddIn.ClientOptions
            );

            var messages = new List<ChatMessage>()
            {
                new SystemChatMessage(
                    "Ты научный редактор. Сформулируй короткий логический переход между двумя фрагментами текста. " +
                    "Переход должен связывать мысль первого фрагмента со вторым и выглядеть естественной частью научного текста. " +
                    "Не переписывай исходные фрагменты, не добавляй новые факты, числа, источники, имена, причинные связи или выводы, которых в них нет. " +
                    "Используй 1–3 предложения. Верни только готовый переходный текст без заголовков и комментариев."
                ),
                new UserChatMessage("Фрагмент 1:\n" + firstFragment),
                new UserChatMessage("Фрагмент 2:\n" + secondFragment),
                new UserChatMessage(
                    "Свяжи эти два фрагмента логически. Не повторяй их содержание дословно и не добавляй неподтверждённых сведений."
                )
            };

            var streamingAnswer = client.CompleteChatStreamingAsync(
                messages,
                new ChatCompletionOptions() { Temperature = 0.14f },
                ThisAddIn.CancellationTokenSource.Token
            );

            StringBuilder response = new StringBuilder();
            CancelButtonVisibility(true);
            SetStatus("◌ Связываю…");

            try
            {
                await foreach (
                    var update in streamingAnswer.WithCancellation(
                        ThisAddIn.CancellationTokenSource.Token
                    )
                )
                {
                    if (ThisAddIn.CancellationTokenSource.IsCancellationRequested)
                        throw new OperationCanceledException();

                    foreach (var newContent in update.ContentUpdate)
                    {
                        if (newContent.Kind == ChatMessageContentPartKind.Text)
                            response.Append(newContent.Text);
                    }
                }
            }
            finally
            {
                CancelButtonVisibility(false);
            }

            return WordMarkdown.RemoveMarkdownSyntax(response.ToString()).Trim();
        }

        private void ResetLogicalLinkSelection()
        {
            _logicalLinkFirstRange = null;
            _logicalLinkFirstText = null;
            _logicalLinkDocumentIdentity = null;

            if (LogicalLinkButton != null)
                LogicalLinkButton.Label = "Связать";
        }

        private static string NormalizeLogicalLinkText(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            return value.Replace("\r", " ").Replace("\n", " ").Trim();
        }

        private static string NormalizeGeneratedTransition(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            string normalized = value.Replace("\r", " ").Replace("\n", " ").Trim();
            while (normalized.Contains("  "))
                normalized = normalized.Replace("  ", " ");

            return normalized;
        }

        private static string GetLogicalLinkDocumentIdentity(Word.Document document)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(document.FullName))
                    return document.FullName;
            }
            catch
            {
            }

            return document.Name;
        }
    }
}
