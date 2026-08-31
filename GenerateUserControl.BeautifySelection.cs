using System;
using System.ClientModel;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenAI.Chat;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    public partial class GenerateUserControl
    {
        private async Task BeautifySelectedWordTextAsync(BeautifyPreset preset)
        {
            if (_beautifyBusy)
                return;

            Word.Selection selection = Globals.ThisAddIn.Application.Selection;
            if (selection == null || selection.End <= selection.Start)
            {
                MessageBox.Show(
                    "Выделите в Word текст, который нужно проанализировать и оформить, затем снова выберите пресет «Сделать красиво».",
                    "НеZнайка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return;
            }

            Word.Range selectedRange = selection.Range.Duplicate;
            string source = (selectedRange.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(source))
            {
                MessageBox.Show(
                    "Выделенный диапазон не содержит текста.",
                    "НеZнайка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return;
            }

            if (ModelProperties.IsImageModel(ThisAddIn.Model))
            {
                MessageBox.Show(
                    "Для анализа и оформления выделенного текста выберите языковую модель.",
                    "НеZнайка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return;
            }

            int sourceTokens = Math.Max(1800, (int)(ThisAddIn.ContextLength * 0.48));
            string boundedSource = CommonUtils.SubstringTokens(source, sourceTokens);
            int allowedDifference = Math.Max(32, source.Length / 100);
            if (boundedSource.Length + allowedDifference < source.Length)
            {
                MessageBox.Show(
                    "Выделение слишком велико для безопасного двухэтапного прохода текущей модели. Уменьшите выделение или обработайте текст несколькими разделами. НеZнайка не будет молча обрезать конец текста.",
                    "НеZнайка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return;
            }

            string previousFinalResponse = _lastResponseMarkdown;
            string previousTemplateName = _lastTemplateName;

            try
            {
                _beautifyBusy = true;
                if (_beautifyButton != null)
                    _beautifyButton.Enabled = false;
                GenerateButton.Enabled = false;
                _insertButton.Enabled = false;
                _copyButton.Enabled = false;

                if (ThisAddIn.CancellationTokenSource == null ||
                    ThisAddIn.CancellationTokenSource.IsCancellationRequested)
                {
                    try { ThisAddIn.CancellationTokenSource?.Dispose(); } catch { }
                    ThisAddIn.CancellationTokenSource = new CancellationTokenSource();
                }

                CancellationToken cancellationToken = ThisAddIn.CancellationTokenSource.Token;
                Forge.CancelButtonVisibility(true);

                ChatClient client = new ChatClient(
                    ThisAddIn.Model,
                    new ApiKeyCredential(ThisAddIn.ApiKey),
                    ThisAddIn.ClientOptions
                );

                Forge.SetModelActivity(true, "Сделать красиво — анализ 1 из 2…");
                if (_responseLabel != null && !_responseLabel.IsDisposed)
                    _responseLabel.Text = "Сделать красиво — анализ выделения 1 из 2…";

                string analysis = await AnalyzeSelectedBeautifyTextAsync(
                    client,
                    boundedSource,
                    cancellationToken
                );
                if (string.IsNullOrWhiteSpace(analysis))
                    throw new InvalidOperationException("Модель не вернула анализ выделенного текста.");

                int analysisTokens = Math.Max(700, (int)(ThisAddIn.ContextLength * 0.10));
                analysis = CommonUtils.SubstringTokens(analysis, analysisTokens);

                string systemPrompt =
                    "Ты научный редактор. Выполни второй этап: преобразуй исходный выделенный текст в выбранный жанр, используя предварительный анализ только как план редактирования. " +
                    "Исходный текст является единственным источником фактов. Не используй внешние знания и не добавляй сведения, которых нет в исходнике. " +
                    "Сохраняй без изменения числа, единицы измерения, формулы, названия веществ и методов, имена, даты, DOI, URL и цитаты. " +
                    "Все ссылки вида [S1], [S2] или [имя.pdf, с. N] сохраняй дословно и рядом с теми утверждениями, которые они подтверждали. " +
                    "Не выдумывай новые ссылки. Если исходник содержит противоречия, неопределенность или разные версии одного утверждения, не сглаживай их без основания. " +
                    "Можно объединять повторяющиеся фрагменты, перестраивать композицию, добавлять логические переходы и улучшать научный стиль. " +
                    "Верни только готовый переработанный текст без комментариев о процессе.";

                string userPrompt =
                    "Формат: " + preset.Name + "\n" +
                    "Инструкция формата: " + preset.Instruction + "\n\n" +
                    "Предварительный анализ структуры (не источник новых фактов):\n<<<ANALYSIS>>>\n" +
                    analysis +
                    "\n<<<END ANALYSIS>>>\n\n" +
                    "Исходный выделенный текст Word:\n<<<SOURCE>>>\n" +
                    boundedSource +
                    "\n<<<END SOURCE>>>";

                Forge.SetModelActivity(true, "Сделать красиво — оформление 2 из 2: " + preset.Name + "…");
                if (_responseLabel != null && !_responseLabel.IsDisposed)
                    _responseLabel.Text = "Сделать красиво — оформление 2 из 2: " + preset.Name + "…";

                _responseTextBox.AppendText(
                    (_responseTextBox.TextLength > 0 ? Environment.NewLine + Environment.NewLine : string.Empty) +
                    "НеZнайка [Сделать красиво → " + preset.Name + "; выделение Word]:" +
                    Environment.NewLine + Environment.NewLine
                );
                ScrollResponseToEnd();

                var answer = client.CompleteChatStreamingAsync(
                    new List<ChatMessage>
                    {
                        new SystemChatMessage(systemPrompt),
                        new UserChatMessage(userPrompt)
                    },
                    new ChatCompletionOptions { Temperature = 0.10f },
                    cancellationToken
                );
                if (answer == null)
                    throw new InvalidOperationException("Модель не вернула поток оформленного текста.");

                string rewritten = await StreamAnswerToPane(answer);
                if (string.IsNullOrWhiteSpace(rewritten))
                    throw new InvalidOperationException("Оформленный ответ модели оказался пустым.");

                List<string> missingCitations = FindMissingProtectedCitations(
                    boundedSource,
                    rewritten
                );
                if (missingCitations.Count > 0)
                {
                    _responseTextBox.AppendText(
                        Environment.NewLine + Environment.NewLine +
                        "[НеZнайка: результат не принят как готовый — модель потеряла ссылки: " +
                        string.Join(", ", missingCitations) + "]"
                    );
                    ScrollResponseToEnd();
                    throw new InvalidOperationException(
                        "При оформлении потеряна одна или несколько исходных ссылок. Выделенный текст Word не изменен."
                    );
                }

                _lastResponseMarkdown = rewritten;
                _lastTemplateName = "Сделать красиво — " + preset.Name;
                _insertButton.Enabled = true;
                _copyButton.Enabled = true;
            }
            catch (OperationCanceledException)
            {
                _lastResponseMarkdown = previousFinalResponse;
                _lastTemplateName = previousTemplateName;
            }
            catch (Exception ex)
            {
                _lastResponseMarkdown = previousFinalResponse;
                _lastTemplateName = previousTemplateName;
                CommonUtils.DisplayWarning(ex);
            }
            finally
            {
                try { Forge.SetModelActivity(false, null); } catch { }
                try { Forge.CancelButtonVisibility(false); } catch { }
                _beautifyBusy = false;
                GenerateButton.Enabled = true;
                if (_beautifyButton != null && !_beautifyButton.IsDisposed)
                    _beautifyButton.Enabled = true;
                _insertButton.Enabled = !string.IsNullOrWhiteSpace(_lastResponseMarkdown);
                _copyButton.Enabled = !string.IsNullOrWhiteSpace(_lastResponseMarkdown);
                if (_responseLabel != null && !_responseLabel.IsDisposed)
                    _responseLabel.Text = "Диалог / ответ:";
                ScrollResponseToEnd();
            }
        }

        private async Task<string> AnalyzeSelectedBeautifyTextAsync(
            ChatClient client,
            string source,
            CancellationToken cancellationToken
        )
        {
            string systemPrompt =
                "Ты анализатор научного текста. Это первый этап редактирования. Не переписывай текст и не используй внешние знания. " +
                "Определи внутреннюю структуру материала, основную тему, смысловые блоки, последовательность аргументов, ключевые факты, повторы, возможные внутренние противоречия, резкие переходы и уже существующие ссылки. " +
                "Отдельно отметь фрагменты, которые нельзя менять по смыслу: числа, формулы, единицы, имена, названия веществ и методов, DOI, URL, цитаты и библиографические ссылки. " +
                "Предложи только план композиционного редактирования. Анализ должен быть компактным и не содержать новых фактов.";

            string userPrompt =
                "Проанализируй выделенный текст перед его последующим оформлением:\n<<<SOURCE>>>\n" +
                source +
                "\n<<<END SOURCE>>>";

            var answer = client.CompleteChatStreamingAsync(
                new List<ChatMessage>
                {
                    new SystemChatMessage(systemPrompt),
                    new UserChatMessage(userPrompt)
                },
                new ChatCompletionOptions { Temperature = 0.03f },
                cancellationToken
            );
            if (answer == null)
                throw new InvalidOperationException("Модель не вернула поток анализа выделенного текста.");

            return await CollectBeautifyStreamAsync(answer, cancellationToken);
        }

        private static async Task<string> CollectBeautifyStreamAsync(
            AsyncCollectionResult<StreamingChatCompletionUpdate> answer,
            CancellationToken cancellationToken
        )
        {
            var result = new StringBuilder();
            await foreach (var update in answer.WithCancellation(cancellationToken))
            {
                foreach (var part in update.ContentUpdate)
                {
                    if (part.Kind == ChatMessageContentPartKind.Text && !string.IsNullOrEmpty(part.Text))
                        result.Append(part.Text);
                }
            }
            return result.ToString();
        }
    }
}
