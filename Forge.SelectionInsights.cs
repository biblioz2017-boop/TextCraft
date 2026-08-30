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
    partial class Forge
    {
        private async void ExplainSelectionButton_Click(object sender, RibbonControlEventArgs e)
        {
            await InsertSelectionInsightAsync(
                "Пояснение",
                "Объясни выделенный фрагмент простым, но научно корректным языком. " +
                "Раскрой смысл терминов, логические связи, причинно-следственные отношения и значение утверждений только в той мере, в какой это следует из выделенного текста. " +
                "Не добавляй внешние факты, новые источники, числа или выводы, которых нет в исходном фрагменте. " +
                "Если в тексте есть неоднозначность или недостаточно данных для однозначного объяснения, явно укажи это. " +
                "Верни связное пояснение без Markdown-заголовка."
            );
        }

        private async void ConclusionsSelectionButton_Click(object sender, RibbonControlEventArgs e)
        {
            await InsertSelectionInsightAsync(
                "Выводы",
                "Сформулируй выводы только по выделенному фрагменту. " +
                "Выдели ключевые результаты, установленные связи, тенденции, ограничения и практическое или научное значение, только если они прямо поддерживаются исходным текстом. " +
                "Не придумывай новые факты, статистику, причинность, источники или обобщения, которых нет в выделении. " +
                "Если текст не позволяет сделать сильный вывод, сформулируй осторожный вывод с соответствующей степенью уверенности. " +
                "Верни 2–6 кратких содержательных выводов, каждый с новой строки, без Markdown-нумерации и без заголовка."
            );
        }

        private static async Task InsertSelectionInsightAsync(string heading, string instruction)
        {
            try
            {
                Word.Selection selection = Globals.ThisAddIn.Application.Selection;
                if (selection == null || selection.End <= selection.Start)
                {
                    MessageBox.Show(
                        "Сначала выделите текст, который нужно обработать.",
                        "НеZнайка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    return;
                }

                Word.Range sourceRange = selection.Range.Duplicate;
                string selectedText = (sourceRange.Text ?? string.Empty).Trim();
                if (selectedText.Length == 0)
                    return;

                SetStatus(heading == "Пояснение" ? "◌ Объясняю выделенное…" : "◌ Формулирую выводы…");
                CancelButtonVisibility(true);

                ChatClient client = new ChatClient(
                    ThisAddIn.Model,
                    new ApiKeyCredential(ThisAddIn.ApiKey),
                    ThisAddIn.ClientOptions
                );

                var messages = new List<ChatMessage>()
                {
                    new SystemChatMessage(
                        "Ты помощник по научному тексту в Microsoft Word. " +
                        "Работай только с предоставленным выделенным фрагментом. " +
                        "Не используй внешние знания как источник новых утверждений. " +
                        "Сохраняй научную точность, числа, единицы измерения, имена, формулы, термины и ссылки."
                    ),
                    new UserChatMessage(instruction + "\n\nВыделенный текст:\n" + selectedText)
                };

                var answer = client.CompleteChatStreamingAsync(
                    messages,
                    new ChatCompletionOptions() { Temperature = 0.10f },
                    ThisAddIn.CancellationTokenSource.Token
                );

                Word.Range insertionRange = sourceRange.Duplicate;
                insertionRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                insertionRange.Text = Environment.NewLine + heading + ":" + Environment.NewLine;
                insertionRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                await AddStreamingChatContentToRange(answer, insertionRange);
                Globals.ThisAddIn.Application.Selection.SetRange(insertionRange.Start, insertionRange.End);
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
                CancelButtonVisibility(false);
                SetStatus("● Готово");
            }
        }
    }
}
