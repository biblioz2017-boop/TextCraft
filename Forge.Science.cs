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
        // Quick rewrite actions use Word's native revision tracking. This gives the user
        // a standard Accept/Reject workflow instead of silently replacing dissertation text.
        private static async Task AnalyzeTextWithTrackChanges(
            string systemPrompt,
            string userPrompt,
            float temperature
        )
        {
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            bool originalTrackRevisions = document.TrackRevisions;

            try
            {
                if (!originalTrackRevisions)
                    document.TrackRevisions = true;

                await AnalyzeText(systemPrompt, userPrompt, temperature);
            }
            finally
            {
                try
                {
                    document.TrackRevisions = originalTrackRevisions;
                }
                catch
                {
                    // Restoring an informational Word setting must not hide the generated text.
                }
            }
        }

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
                    "НеZнайка",
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

        private async void TranslateLanguageButton_Click(object sender, RibbonControlEventArgs e)
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
                        "НеZнайка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    return;
                }

                await AnalyzeTextWithTrackChanges(QuickTextSystemPrompt, instruction, temperature);
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
    }
}
