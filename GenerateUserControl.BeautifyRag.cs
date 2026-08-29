using System;
using System.ClientModel;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenAI.Chat;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    public partial class GenerateUserControl
    {
        private CheckBox _beautifyUseRagCheckBox;
        private ToolTip _beautifyUseRagToolTip;

        private sealed class BeautifyGap
        {
            public int Number { get; set; }
            public string Title { get; set; }
            public string Need { get; set; }
            public string Query { get; set; }
        }

        private sealed class BeautifyRagSource
        {
            public int MarkerIndex { get; set; }
            public RAGControl.RagEvidenceItem Evidence { get; set; }
        }

        private sealed class BeautifyGapResult
        {
            public BeautifyGap Gap { get; set; }
            public List<BeautifyRagSource> Sources { get; private set; }

            public BeautifyGapResult()
            {
                Sources = new List<BeautifyRagSource>();
            }
        }

        private void InitializeBeautifyRagOption()
        {
            if (_beautifyUseRagCheckBox != null)
                return;

            _beautifyUseRagCheckBox = new CheckBox
            {
                Text = "Дополнить из RAG",
                AutoSize = true,
                Checked = false,
                Height = 26,
                Margin = new Padding(4, 3, 6, 2)
            };

            _beautifyUseRagToolTip = new ToolTip();
            _beautifyUseRagToolTip.SetToolTip(
                _beautifyUseRagCheckBox,
                "После анализа выделенного текста НеZнайка найдет содержательные пробелы, " +
                "поищет подтверждения только в отмеченных PDF и заполнит только подтвержденные пробелы."
            );
        }

        private async Task BeautifySelectedWordTextWithOptionalRagAsync(BeautifyPreset preset)
        {
            bool useRag = _beautifyUseRagCheckBox != null && _beautifyUseRagCheckBox.Checked;
            if (!useRag)
            {
                await BeautifySelectedWordTextAsync(preset);
                return;
            }

            await BeautifySelectedWordTextWithRagAsync(preset);
        }

        private async Task BeautifySelectedWordTextWithRagAsync(BeautifyPreset preset)
        {
            if (_beautifyBusy)
                return;

            Word.Selection selection = Globals.ThisAddIn.Application.Selection;
            if (selection == null || selection.End <= selection.Start)
            {
                MessageBox.Show(
                    "Выделите в Word текст, который нужно проанализировать, дополнить из RAG и оформить.",
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
                    "Для анализа, RAG-поиска и оформления выберите языковую модель.",
                    "НеZнайка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return;
            }

            RAGControl rag;
            if (!TryGetRagControlForActiveDocument(out rag) || rag == null)
            {
                MessageBox.Show(
                    "Не удалось связать «Сделать красиво» с панелью «Литература» текущего документа. Откройте панель литературы и повторите попытку.",
                    "НеZнайка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return;
            }

            int sourceTokens = Math.Max(1600, (int)(ThisAddIn.ContextLength * 0.34));
            string boundedSource = CommonUtils.SubstringTokens(source, sourceTokens);
            int allowedDifference = Math.Max(32, source.Length / 100);
            if (boundedSource.Length + allowedDifference < source.Length)
            {
                MessageBox.Show(
                    "Выделение слишком велико для безопасного режима «Сделать красиво + RAG». Уменьшите выделение или обработайте материал несколькими разделами. НеZнайка не будет молча обрезать текст.",
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
                if (_beautifyUseRagCheckBox != null)
                    _beautifyUseRagCheckBox.Enabled = false;
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

                Forge.SetModelActivity(true, "Сделать красиво + RAG — анализ 1 из 4…");
                SetBeautifyStatus("Сделать красиво + RAG — анализ выделения 1 из 4…");

                string analysis = await AnalyzeBeautifyGapsAsync(
                    client,
                    boundedSource,
                    preset,
                    cancellationToken
                );
                if (string.IsNullOrWhiteSpace(analysis))
                    throw new InvalidOperationException("Модель не вернула анализ выделенного текста.");

                int analysisTokens = Math.Max(900, (int)(ThisAddIn.ContextLength * 0.12));
                analysis = CommonUtils.SubstringTokens(analysis, analysisTokens);
                List<BeautifyGap> gaps = ParseBeautifyGaps(analysis);

                Forge.SetModelActivity(true, "Сделать красиво + RAG — поиск 2 из 4…");
                SetBeautifyStatus(
                    gaps.Count == 0
                        ? "Сделать красиво + RAG — 2 из 4: содержательных пробелов не найдено."
                        : "Сделать красиво + RAG — поиск материалов 2 из 4: 0/" + gaps.Count
                );

                int originalMaxMarker = GetMaxSourceMarkerIndex(boundedSource);
                int nextMarker = originalMaxMarker + 1;
                List<BeautifyGapResult> gapResults = new List<BeautifyGapResult>();
                List<BeautifyRagSource> allSources = new List<BeautifyRagSource>();
                Dictionary<string, BeautifyRagSource> uniqueSources =
                    new Dictionary<string, BeautifyRagSource>(StringComparer.Ordinal);

                const int maxTotalSources = 10;
                const int maxSourcesPerGap = 3;

                for (int gapIndex = 0; gapIndex < gaps.Count; gapIndex++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    BeautifyGap gap = gaps[gapIndex];
                    SetBeautifyStatus(
                        "Сделать красиво + RAG — поиск материалов 2 из 4: " +
                        (gapIndex + 1) + "/" + gaps.Count + " — " + gap.Title
                    );

                    BeautifyGapResult gapResult = new BeautifyGapResult { Gap = gap };
                    gapResults.Add(gapResult);

                    if (allSources.Count >= maxTotalSources || string.IsNullOrWhiteSpace(gap.Query))
                        continue;

                    List<RAGControl.RagEvidenceItem> found = await Task.Run(
                        () => rag.GetRAGEvidence(gap.Query, 2),
                        cancellationToken
                    );

                    int acceptedForGap = 0;
                    foreach (RAGControl.RagEvidenceItem item in found)
                    {
                        if (item == null || string.IsNullOrWhiteSpace(item.Text))
                            continue;

                        string key = BuildBeautifyEvidenceKey(item);
                        BeautifyRagSource sourceBinding;
                        if (!uniqueSources.TryGetValue(key, out sourceBinding))
                        {
                            if (allSources.Count >= maxTotalSources)
                                break;

                            sourceBinding = new BeautifyRagSource
                            {
                                MarkerIndex = nextMarker++,
                                Evidence = item
                            };
                            uniqueSources[key] = sourceBinding;
                            allSources.Add(sourceBinding);
                        }

                        if (!gapResult.Sources.Contains(sourceBinding))
                        {
                            gapResult.Sources.Add(sourceBinding);
                            acceptedForGap++;
                        }

                        if (acceptedForGap >= maxSourcesPerGap)
                            break;
                    }
                }

                RenderBeautifyRagEvidence(gapResults);
                string evidenceBlock = BuildBeautifyRagEvidenceBlock(gapResults);

                Forge.SetModelActivity(true, "Сделать красиво + RAG — оформление 3 из 4…");
                SetBeautifyStatus("Сделать красиво + RAG — формирую " + preset.Name + " — 3 из 4…");

                string systemPrompt =
                    "Ты научный редактор, работающий в строгом режиме дополнения из RAG. " +
                    "Исходный выделенный текст Word является основной рукописью. Предварительный анализ задает только план и список пробелов, но не является источником фактов. " +
                    "Новые факты разрешено добавлять исключительно из блока ПРОВЕРЕННЫЕ RAG-ФРАГМЕНТЫ. Не используй внешние знания. " +
                    "Каждое добавленное из RAG фактическое положение обязательно сопровождай соответствующей меткой [S#] из предоставленного фрагмента. " +
                    "Если для пробела написано НЕТ ПОДТВЕРЖДАЮЩИХ ФРАГМЕНТОВ, не заполняй этот пробел догадками. " +
                    "Сохраняй исходные числа, единицы измерения, формулы, имена, названия веществ и методов, даты, DOI, URL, цитаты и уже существующие ссылки. " +
                    "Не создавай новые библиографические записи и не меняй смысл исходных утверждений. " +
                    "Можно устранять повторы, перестраивать композицию, объединять близкие фрагменты и добавлять логические переходы. " +
                    "Верни только готовый текст выбранного жанра; не выводи служебные GAP-блоки, план анализа или комментарии о процессе.";

                string userPrompt =
                    "Формат: " + preset.Name + "\n" +
                    "Инструкция формата: " + preset.Instruction + "\n\n" +
                    "АНАЛИЗ И КАРТА ПРОБЕЛОВ:\n<<<ANALYSIS>>>\n" +
                    analysis +
                    "\n<<<END ANALYSIS>>>\n\n" +
                    evidenceBlock +
                    "\n\nИСХОДНЫЙ ВЫДЕЛЕННЫЙ ТЕКСТ WORD:\n<<<SOURCE>>>\n" +
                    boundedSource +
                    "\n<<<END SOURCE>>>";

                _responseTextBox.AppendText(
                    (_responseTextBox.TextLength > 0 ? Environment.NewLine + Environment.NewLine : string.Empty) +
                    "НеZнайка [Сделать красиво → " + preset.Name + "; выделение Word + RAG]:" +
                    Environment.NewLine + Environment.NewLine
                );
                int responseStart = _responseTextBox.TextLength;
                ScrollResponseToEnd();

                var answer = client.CompleteChatStreamingAsync(
                    new List<ChatMessage>
                    {
                        new SystemChatMessage(systemPrompt),
                        new UserChatMessage(userPrompt)
                    },
                    new ChatCompletionOptions { Temperature = 0.08f },
                    cancellationToken
                );
                if (answer == null)
                    throw new InvalidOperationException("Модель не вернула поток итогового текста.");

                string rewritten = await StreamAnswerToPane(answer);
                if (string.IsNullOrWhiteSpace(rewritten))
                    throw new InvalidOperationException("Итоговый ответ модели оказался пустым.");

                Forge.SetModelActivity(true, "Сделать красиво + RAG — проверка 4 из 4…");
                SetBeautifyStatus("Сделать красиво + RAG — проверяю ссылки и добавленные факты — 4 из 4…");

                HashSet<int> usedRagMarkers;
                string grounded = GroundBeautifyRagResponse(
                    rewritten,
                    allSources,
                    originalMaxMarker,
                    out usedRagMarkers
                );

                List<string> missingCitations = FindMissingProtectedCitations(
                    boundedSource,
                    grounded
                );
                if (missingCitations.Count > 0)
                {
                    throw new InvalidOperationException(
                        "При переработке потеряны исходные ссылки: " +
                        string.Join(", ", missingCitations) + ". Выделенный текст Word не изменен."
                    );
                }

                ReplaceBeautifyResponse(responseStart, grounded);
                _lastResponseMarkdown = grounded;
                _lastTemplateName = "Сделать красиво + RAG — " + preset.Name;
                _insertButton.Enabled = true;
                _copyButton.Enabled = true;

                AppendBeautifyRagSummary(gapResults, allSources, usedRagMarkers);
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
                if (_beautifyUseRagCheckBox != null && !_beautifyUseRagCheckBox.IsDisposed)
                    _beautifyUseRagCheckBox.Enabled = true;
                _insertButton.Enabled = !string.IsNullOrWhiteSpace(_lastResponseMarkdown);
                _copyButton.Enabled = !string.IsNullOrWhiteSpace(_lastResponseMarkdown);
                SetBeautifyStatus("Диалог / ответ:");
                ScrollResponseToEnd();
            }
        }

        private async Task<string> AnalyzeBeautifyGapsAsync(
            ChatClient client,
            string source,
            BeautifyPreset preset,
            CancellationToken cancellationToken
        )
        {
            string systemPrompt =
                "Ты анализатор научной рукописи. Не переписывай текст и не используй внешние знания. " +
                "Определи композицию, уже раскрытые вопросы, повторы, логические разрывы и содержательные пробелы относительно выбранного жанра. " +
                "Содержательным пробелом считай только то, для чего действительно нужно найти дополнительный научный материал; не требуй сведений, не обязательных для данного текста. " +
                "Для каждого пробела сформируй короткий самостоятельный поисковый запрос, пригодный для семантического поиска по научным PDF. " +
                "Не более пяти пробелов. Ответ должен использовать точный формат блоков ниже. Если содержательных пробелов нет, напиши NO_GAPS.\n\n" +
                "Краткий план композиционного редактирования.\n" +
                "<<<GAP>>>\nTITLE: краткое название пробела\nNEED: какие сведения нужны\nQUERY: конкретный поисковый запрос по существам/механизмам/методам/результатам\n<<<END GAP>>>";

            string userPrompt =
                "Целевой формат: " + preset.Name + "\n" +
                "Требования формата: " + preset.Instruction + "\n\n" +
                "Проанализируй выделенный текст:\n<<<SOURCE>>>\n" +
                source +
                "\n<<<END SOURCE>>>";

            var answer = client.CompleteChatStreamingAsync(
                new List<ChatMessage>
                {
                    new SystemChatMessage(systemPrompt),
                    new UserChatMessage(userPrompt)
                },
                new ChatCompletionOptions { Temperature = 0.02f },
                cancellationToken
            );
            if (answer == null)
                throw new InvalidOperationException("Модель не вернула поток анализа пробелов.");

            return await CollectBeautifyStreamAsync(answer, cancellationToken);
        }

        private static List<BeautifyGap> ParseBeautifyGaps(string analysis)
        {
            List<BeautifyGap> gaps = new List<BeautifyGap>();
            if (string.IsNullOrWhiteSpace(analysis) ||
                analysis.IndexOf("NO_GAPS", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return gaps;
            }

            MatchCollection blocks = Regex.Matches(
                analysis,
                @"<<<GAP>>>\s*(.*?)\s*<<<END GAP>>>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase
            );

            foreach (Match block in blocks)
            {
                if (gaps.Count >= 5)
                    break;

                string body = block.Groups[1].Value;
                string title = ExtractBeautifyGapField(body, "TITLE");
                string need = ExtractBeautifyGapField(body, "NEED");
                string query = ExtractBeautifyGapField(body, "QUERY");

                if (string.IsNullOrWhiteSpace(query))
                    query = (title + " " + need).Trim();
                if (string.IsNullOrWhiteSpace(query))
                    continue;

                gaps.Add(
                    new BeautifyGap
                    {
                        Number = gaps.Count + 1,
                        Title = string.IsNullOrWhiteSpace(title) ? "Пробел " + (gaps.Count + 1) : title,
                        Need = need,
                        Query = query
                    }
                );
            }

            return gaps;
        }

        private static string ExtractBeautifyGapField(string body, string field)
        {
            Match match = Regex.Match(
                body ?? string.Empty,
                @"(?im)^\s*" + Regex.Escape(field) + @"\s*:\s*(.+?)\s*$"
            );
            return match.Success ? match.Groups[1].Value.Trim() : string.Empty;
        }

        private static int GetMaxSourceMarkerIndex(string text)
        {
            int max = 0;
            foreach (Match match in Regex.Matches(text ?? string.Empty, @"\[S(\d+)\]", RegexOptions.IgnoreCase))
            {
                int parsed;
                if (int.TryParse(match.Groups[1].Value, out parsed) && parsed > max)
                    max = parsed;
            }
            return max;
        }

        private static string BuildBeautifyEvidenceKey(RAGControl.RagEvidenceItem item)
        {
            return (item.Source ?? string.Empty) + "|" +
                (item.Page.HasValue ? item.Page.Value.ToString() : string.Empty) + "|" +
                (item.Text ?? string.Empty);
        }

        private string BuildBeautifyRagEvidenceBlock(List<BeautifyGapResult> gapResults)
        {
            StringBuilder block = new StringBuilder();
            block.AppendLine("ПРОВЕРЕННЫЕ RAG-ФРАГМЕНТЫ ДЛЯ ЗАПОЛНЕНИЯ ПРОБЕЛОВ:");
            block.AppendLine("Используй фрагменты только для соответствующих пробелов. Метки [S#] обязательны для добавленных фактов.");

            int evidenceTokens = Math.Max(260, (int)(ThisAddIn.ContextLength * 0.015));
            foreach (BeautifyGapResult result in gapResults)
            {
                block.AppendLine();
                block.Append("GAP ").Append(result.Gap.Number).Append(": ").AppendLine(result.Gap.Title);
                if (!string.IsNullOrWhiteSpace(result.Gap.Need))
                    block.Append("Нужно: ").AppendLine(result.Gap.Need);

                if (result.Sources.Count == 0)
                {
                    block.AppendLine("НЕТ ПОДТВЕРЖДАЮЩИХ ФРАГМЕНТОВ — НЕ ЗАПОЛНЯТЬ ЭТОТ ПРОБЕЛ НОВЫМИ ФАКТАМИ.");
                    continue;
                }

                foreach (BeautifyRagSource source in result.Sources)
                {
                    RAGControl.RagEvidenceItem item = source.Evidence;
                    block.Append("[S").Append(source.MarkerIndex).Append("] ")
                        .Append(item.Source ?? "источник");
                    if (item.Page.HasValue)
                        block.Append(", с. ").Append(item.Page.Value);
                    block.AppendLine();
                    block.AppendLine(CommonUtils.SubstringTokens(item.Text ?? string.Empty, evidenceTokens));
                }
            }

            return block.ToString();
        }

        private void RenderBeautifyRagEvidence(List<BeautifyGapResult> gapResults)
        {
            if (_evidenceTextBox == null || _evidenceTextBox.IsDisposed)
                return;

            StringBuilder text = new StringBuilder();
            foreach (BeautifyGapResult result in gapResults)
            {
                text.Append("Пробел ").Append(result.Gap.Number).Append(": ").AppendLine(result.Gap.Title);
                if (result.Sources.Count == 0)
                {
                    text.AppendLine("  подтверждающих фрагментов не найдено");
                }
                else
                {
                    foreach (BeautifyRagSource source in result.Sources)
                    {
                        text.Append("  [S").Append(source.MarkerIndex).Append("] ")
                            .AppendLine(source.Evidence.CitationLabel);
                    }
                }
                text.AppendLine();
            }

            _evidenceTextBox.Text = text.ToString().TrimEnd();
        }

        private static string GroundBeautifyRagResponse(
            string response,
            List<BeautifyRagSource> sources,
            int originalMaxMarker,
            out HashSet<int> usedMarkers
        )
        {
            usedMarkers = new HashSet<int>();
            Dictionary<int, BeautifyRagSource> known = new Dictionary<int, BeautifyRagSource>();
            foreach (BeautifyRagSource source in sources)
                known[source.MarkerIndex] = source;

            MatchCollection matches = Regex.Matches(
                response ?? string.Empty,
                @"\[S(\d+)\]",
                RegexOptions.IgnoreCase
            );

            foreach (Match match in matches)
            {
                int marker;
                if (!int.TryParse(match.Groups[1].Value, out marker))
                    continue;

                if (marker <= originalMaxMarker)
                    continue;

                if (!known.ContainsKey(marker))
                {
                    throw new InvalidOperationException(
                        "Модель сослалась на отсутствующий RAG-фрагмент [S" + marker + "]. Результат не принят как готовый."
                    );
                }
                usedMarkers.Add(marker);
            }

            if (sources.Count > 0 && usedMarkers.Count == 0)
            {
                throw new InvalidOperationException(
                    "RAG нашел дополнительные материалы, но итоговый текст не связал добавленные сведения ни с одной меткой [S#]. Результат не принят как готовый."
                );
            }

            return Regex.Replace(
                response ?? string.Empty,
                @"\[S(\d+)\]",
                delegate(Match match)
                {
                    int marker;
                    if (!int.TryParse(match.Groups[1].Value, out marker))
                        return match.Value;
                    if (marker <= originalMaxMarker)
                        return match.Value;

                    BeautifyRagSource source;
                    if (!known.TryGetValue(marker, out source))
                        return match.Value;
                    return source.Evidence.CitationLabel;
                },
                RegexOptions.IgnoreCase
            );
        }

        private void ReplaceBeautifyResponse(int responseStart, string response)
        {
            if (_responseTextBox == null || _responseTextBox.IsDisposed)
                return;

            responseStart = Math.Max(0, Math.Min(responseStart, _responseTextBox.TextLength));
            _responseTextBox.Select(responseStart, _responseTextBox.TextLength - responseStart);
            _responseTextBox.SelectedText = response ?? string.Empty;
            _responseTextBox.SelectionStart = _responseTextBox.TextLength;
            _responseTextBox.SelectionLength = 0;
            ScrollResponseToEnd();
        }

        private void AppendBeautifyRagSummary(
            List<BeautifyGapResult> gapResults,
            List<BeautifyRagSource> allSources,
            HashSet<int> usedMarkers
        )
        {
            int withEvidence = 0;
            int withoutEvidence = 0;
            int filled = 0;

            foreach (BeautifyGapResult result in gapResults)
            {
                if (result.Sources.Count == 0)
                {
                    withoutEvidence++;
                    continue;
                }

                withEvidence++;
                bool used = false;
                foreach (BeautifyRagSource source in result.Sources)
                {
                    if (usedMarkers.Contains(source.MarkerIndex))
                    {
                        used = true;
                        break;
                    }
                }
                if (used)
                    filled++;
            }

            _responseTextBox.AppendText(
                Environment.NewLine + Environment.NewLine +
                "[НеZнайка — RAG-дополнение] Пробелов: " + gapResults.Count +
                "; с найденными материалами: " + withEvidence +
                "; использовано при дополнении: " + filled +
                "; без подтверждающих материалов: " + withoutEvidence +
                "; уникальных RAG-фрагментов: " + allSources.Count + "."
            );
            ScrollResponseToEnd();
        }

        private void SetBeautifyStatus(string text)
        {
            if (_responseLabel == null || _responseLabel.IsDisposed)
                return;
            try { _responseLabel.Text = text ?? string.Empty; } catch { }
        }
    }
}
