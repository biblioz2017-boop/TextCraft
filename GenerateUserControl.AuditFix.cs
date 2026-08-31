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
        private Button _fixAuditButton;
        private Button _nextAuditFixButton;
        private Button _resetAuditButton;
        private ToolTip _auditFixToolTip;
        private Word.Range _auditTargetRange;
        private string _auditTargetDocumentIdentity;
        private string _lastAuditReport = string.Empty;
        private readonly List<string> _appliedAuditFixes = new List<string>();
        private bool _auditFixBusy;

        private sealed class AuditEdit
        {
            public string FindText { get; set; }
            public string Replacement { get; set; }
            public string Reason { get; set; }
        }

        private sealed class ResolvedAuditEdit
        {
            public AuditEdit Edit { get; set; }
            public int RelativeStart { get; set; }
        }

        private void InitializeAuditFixControls()
        {
            if (_fixAuditButton != null)
                return;

            _fixAuditButton = new Button
            {
                Text = "Исправить аудит",
                AutoSize = true,
                Height = 28,
                Enabled = false
            };
            _fixAuditButton.Click += FixAuditButton_Click;

            _nextAuditFixButton = new Button
            {
                Text = "След. правка",
                AutoSize = true,
                Height = 28,
                Enabled = false
            };
            _nextAuditFixButton.Click += NextAuditFixButton_Click;

            _resetAuditButton = new Button
            {
                Text = "Сброс аудита",
                AutoSize = true,
                Height = 28,
                Enabled = false
            };
            _resetAuditButton.Click += ResetAuditButton_Click;

            _auditFixToolTip = new ToolTip();
            _auditFixToolTip.SetToolTip(
                _fixAuditButton,
                "Применить все безопасные редакторские исправления из последнего научного аудита. " +
                "Правки вносятся через рецензирование Word. Факты, числа, ссылки и спорные утверждения автоматически не изменяются."
            );
            _auditFixToolTip.SetToolTip(
                _nextAuditFixButton,
                "Применить только одну следующую безопасную правку из аудита. " +
                "Удобно для последовательной проверки изменений."
            );
            _auditFixToolTip.SetToolTip(
                _resetAuditButton,
                "Забыть сохраненный аудит и привязанный к нему фрагмент документа."
            );

            _quickActionsPanel.WrapContents = true;
            _quickActionsPanel.Controls.Add(_fixAuditButton);
            _quickActionsPanel.Controls.Add(_nextAuditFixButton);
            _quickActionsPanel.Controls.Add(_resetAuditButton);

            // The narrow Word task pane needs four wrapped rows after the RAG checkbox
            // and the existing scientific quick actions are present.
            if (_mainLayout != null && _mainLayout.RowStyles.Count > 2)
                _mainLayout.RowStyles[2].Height = Math.Max(_mainLayout.RowStyles[2].Height, 128F);

            if (_clearButton != null)
                _clearButton.Click += (s, e) => ResetAuditFixState(false);
        }

        private void RememberAuditTarget(Word.Range selectedRange, string auditedText)
        {
            ResetAuditFixState(false);

            if (selectedRange == null || selectedRange.End <= selectedRange.Start)
                return;

            Word.Document document = selectedRange.Document;
            string selectedText = selectedRange.Text ?? string.Empty;
            string effectiveAuditText = auditedText ?? string.Empty;

            int relativeStart = 0;
            int effectiveLength = selectedText.Length;

            if (!string.IsNullOrEmpty(effectiveAuditText))
            {
                int exact = selectedText.IndexOf(effectiveAuditText, StringComparison.Ordinal);
                if (exact >= 0)
                {
                    relativeStart = exact;
                    effectiveLength = effectiveAuditText.Length;
                }
                else
                {
                    // SubstringTokens normally returns a prefix. If whitespace trimming
                    // prevents an exact match, keep the audited length rather than later
                    // rewriting text that the model never actually audited.
                    string probe = effectiveAuditText.Length > 80
                        ? effectiveAuditText.Substring(0, 80)
                        : effectiveAuditText;
                    int probeStart = selectedText.IndexOf(probe, StringComparison.Ordinal);
                    if (probeStart >= 0)
                        relativeStart = probeStart;

                    effectiveLength = Math.Min(effectiveAuditText.Length, selectedText.Length - relativeStart);
                }
            }

            int start = selectedRange.Start + relativeStart;
            int end = Math.Min(selectedRange.End, start + Math.Max(0, effectiveLength));
            if (end <= start)
                return;

            _auditTargetRange = document.Range(start, end);
            _auditTargetDocumentIdentity = GetAuditDocumentIdentity(document);
            _appliedAuditFixes.Clear();
            SetAuditFixButtons(false);
        }

        private void PrepareAuditTargetForRequest(
            string templateName,
            Word.Range requestRange,
            string auditedText
        )
        {
            if (!string.Equals(templateName, "Научный аудит", StringComparison.Ordinal))
                return;

            if (requestRange != null && requestRange.End > requestRange.Start)
            {
                string selectionText = (requestRange.Text ?? string.Empty).Trim();
                string effectiveText = (auditedText ?? string.Empty).Trim();

                // A manually entered audit request may not literally equal the selection.
                // In that case bind the report to the whole selected range instead of
                // silently keeping a stale target or an arbitrary prefix.
                if (effectiveText.Length == 0 ||
                    selectionText.IndexOf(effectiveText, StringComparison.Ordinal) < 0)
                {
                    effectiveText = selectionText;
                }

                RememberAuditTarget(requestRange.Duplicate, effectiveText);
            }
            else
            {
                ResetAuditFixState(false);
            }

            PrepareAuditReviewForNewRun();
        }

        private void CaptureAuditResultIfNeeded(string templateName, string response)
        {
            if (!string.Equals(templateName, "Научный аудит", StringComparison.Ordinal))
                return;

            if (_auditTargetRange == null)
                return;

            string report = (response ?? string.Empty).Trim();
            if (report.Length == 0)
                return;

            if (report.StartsWith("Строгий RAG остановил генерацию", StringComparison.OrdinalIgnoreCase) ||
                report.StartsWith("⚠ Строгий RAG отклонил ответ", StringComparison.OrdinalIgnoreCase))
            {
                SetAuditFixButtons(false);
                return;
            }

            _lastAuditReport = report;
            _appliedAuditFixes.Clear();
            SetAuditFixButtons(true);
        }

        private async void FixAuditButton_Click(object sender, EventArgs e)
        {
            await RunAuditFixAsync(false);
        }

        private async void NextAuditFixButton_Click(object sender, EventArgs e)
        {
            await RunAuditFixAsync(true);
        }

        private void ResetAuditButton_Click(object sender, EventArgs e)
        {
            ResetAuditFixState(true);
        }

        private static void EnsureLocalAuditEndpoint()
        {
            Uri endpoint;
            if (!Uri.TryCreate(ThisAddIn.OpenAIEndpoint, UriKind.Absolute, out endpoint))
            {
                throw new InvalidOperationException(
                    "Научный аудит остановлен: адрес LLM endpoint задан некорректно."
                );
            }

            string host = (endpoint.Host ?? string.Empty).Trim().Trim('[', ']');
            bool isLoopback =
                host.Equals("localhost", StringComparison.OrdinalIgnoreCase) ||
                host.Equals("127.0.0.1", StringComparison.OrdinalIgnoreCase) ||
                host.Equals("::1", StringComparison.OrdinalIgnoreCase);

            if (isLoopback)
                return;

            throw new InvalidOperationException(
                "Научный аудит разрешен только через локальный LLM endpoint " +
                "(localhost, 127.0.0.1 или ::1). Текущий адрес: " +
                ThisAddIn.OpenAIEndpoint +
                ". Измените TEXTCRAFT_OPENAI_ENDPOINT и перезапустите Word."
            );
        }

        private static CancellationToken GetAuditOperationToken()
        {
            CancellationTokenSource source = ThisAddIn.CancellationTokenSource;
            if (source == null || source.IsCancellationRequested)
            {
                source = new CancellationTokenSource();
                ThisAddIn.CancellationTokenSource = source;
            }

            return source.Token;
        }

        private async Task RunAuditFixAsync(bool singleEdit)
        {
            if (_auditFixBusy)
                return;

            if (_auditReviewBusy)
            {
                MessageBox.Show(
                    "Дождитесь, пока НеZнайка закончит разбирать отчет на отдельные замечания.",
                    "НеZнайка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return;
            }

            Word.Document document;
            Word.Range targetRange;
            if (!TryGetCurrentAuditTarget(out document, out targetRange))
                return;

            if (ModelProperties.IsImageModel(ThisAddIn.Model))
            {
                MessageBox.Show(
                    "Для исправления по аудиту выберите языковую модель.",
                    "НеZнайка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return;
            }

            string currentText = targetRange.Text ?? string.Empty;
            if (string.IsNullOrWhiteSpace(currentText))
            {
                MessageBox.Show(
                    "Фрагмент, к которому относился аудит, больше не содержит текста.",
                    "НеZнайка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                ResetAuditFixState(false);
                return;
            }

            string stage = "подготовка";
            bool modelActivityStarted = false;
            try
            {
                _auditFixBusy = true;
                SetAuditControlsBusy(true);

                int maxEdits = singleEdit ? 1 : 16;
                List<AuditEdit> edits = GetPendingSafeAuditReviewEdits(maxEdits);
                bool usingPreparedIssues = edits.Count > 0;

                if (usingPreparedIssues)
                {
                    if (_responseLabel != null)
                        _responseLabel.Text = singleEdit
                            ? "Аудит — применяю следующую проверенную правку…"
                            : "Аудит — применяю отмеченные безопасные правки…";
                }
                else
                {
                    stage = "запрос безопасных правок у модели";
                    Forge.SetModelActivity(
                        true,
                        singleEdit ? "Готовит следующую правку…" : "Готовит правки аудита…"
                    );
                    modelActivityStarted = true;
                    if (_responseLabel != null)
                        _responseLabel.Text = singleEdit
                            ? "Аудит — готовлю следующую безопасную правку…"
                            : "Аудит — готовлю безопасные правки…";

                    edits = await GenerateAuditEditsAsync(currentText, _lastAuditReport, maxEdits);
                }

                if (edits == null || edits.Count == 0)
                {
                    AppendAuditFixNotice(
                        singleEdit
                            ? "Безопасных неиспользованных правок по этому аудиту больше не найдено."
                            : "Автоматически безопасные правки по этому аудиту не найдены. Замечания о фактах, источниках и противоречиях оставлены для ручной проверки."
                    );
                    return;
                }

                stage = "проверка и применение правок в Word";
                int skipped;
                List<AuditEdit> appliedEdits;
                int applied = ApplyAuditEdits(document, targetRange, edits, out skipped, out appliedEdits);

                if (usingPreparedIssues && appliedEdits.Count > 0)
                    MarkAuditReviewEditsApplied(appliedEdits);

                if (applied == 0)
                {
                    AppendAuditFixNotice(
                        "Модель предложила правки, но НеZнайка не применила их: исходный фрагмент не совпал однозначно либо правка затрагивала защищенные числа/ссылки."
                    );
                    return;
                }

                AppendAuditFixNotice(
                    "Применено правок через рецензирование Word: " + applied +
                    (skipped > 0 ? ". Пропущено небезопасных или неоднозначных: " + skipped + "." : ".") +
                    " Изменения можно принять или отклонить стандартными средствами Word."
                );
            }
            catch (OperationCanceledException)
            {
                AppendAuditFixNotice("Исправление по аудиту остановлено пользователем.");
            }
            catch (Exception ex)
            {
                string context = "Исправление аудита остановлено на этапе «" + stage + "»";
                AppendAuditFixNotice(context + ". Проверьте уже внесенные изменения в режиме рецензирования Word.");
                CommonUtils.DisplayError(context, ex);
            }
            finally
            {
                if (modelActivityStarted)
                    Forge.SetModelActivity(false, null);
                _auditFixBusy = false;
                SetAuditControlsBusy(false);
                if (_responseLabel != null)
                    _responseLabel.Text = "Диалог / ответ:";
            }
        }

        private async Task<List<AuditEdit>> GenerateAuditEditsAsync(
            string currentText,
            string auditReport,
            int maxEdits
        )
        {
            EnsureLocalAuditEndpoint();

            int textTokens = Math.Max(1200, (int)(ThisAddIn.ContextLength * 0.38));
            int auditTokens = Math.Max(800, (int)(ThisAddIn.ContextLength * 0.24));

            string boundedText = CommonUtils.SubstringTokens(currentText, textTokens);
            string boundedAudit = CommonUtils.SubstringTokens(auditReport ?? string.Empty, auditTokens);

            StringBuilder applied = new StringBuilder();
            if (_appliedAuditFixes.Count > 0)
            {
                applied.AppendLine("Уже примененные правки, которые нельзя предлагать повторно:");
                foreach (string item in _appliedAuditFixes)
                    applied.AppendLine("- " + item);
            }

            string systemPrompt =
                "Ты научный редактор диссертации. Преобразуй диагностический аудит в точечные безопасные редакторские правки. " +
                "Автоматически разрешено исправлять только: повторы, громоздкий синтаксис, неясную связь предложений, " +
                "расплывчатые или оценочные формулировки, очевидную терминологическую непоследовательность и языковые ошибки. " +
                "Запрещено автоматически: добавлять или удалять факты; менять числа, единицы измерения, имена, даты, формулы, цитаты, DOI, URL и библиографические ссылки; " +
                "подбирать отсутствующие источники; разрешать содержательные противоречия; усиливать или ослаблять научный вывод; " +
                "раскрывать аббревиатуру, если расшифровка явно не присутствует в исходном тексте. " +
                "Если безопасной правки нет, верни только <none/>. " +
                "Для каждой правки FIND должен быть дословной уникальной подстрокой текущего текста, желательно в пределах одного предложения. " +
                "Не используй Markdown и не добавляй текст вне заданных тегов.";

            string userPrompt =
                "Научный аудит:\n" + boundedAudit + "\n\n" +
                "Текущий текст после уже выполненных правок:\n<<<TEXT>>>\n" + boundedText + "\n<<<END TEXT>>>\n\n" +
                applied.ToString() + "\n" +
                "Верни не более " + maxEdits + " безопасных правок строго в формате:\n" +
                "<edit>\n<find>дословный фрагмент исходного текста</find>\n<replace>исправленный фрагмент</replace>\n<reason>краткая причина</reason>\n</edit>\n" +
                "Если замечание требует источника, фактического решения или изменения числа/ссылки, не превращай его в edit.";

            CancellationToken cancellationToken = GetAuditOperationToken();

            ChatClient client = new ChatClient(
                ThisAddIn.Model,
                new ApiKeyCredential(ThisAddIn.ApiKey),
                ThisAddIn.ClientOptions
            );

            var messages = new List<ChatMessage>
            {
                new SystemChatMessage(systemPrompt),
                new UserChatMessage(userPrompt)
            };

            var answer = client.CompleteChatStreamingAsync(
                messages,
                new ChatCompletionOptions { Temperature = 0.05f },
                cancellationToken
            );
            if (answer == null)
                throw new InvalidOperationException("Модель не вернула поток безопасных правок.");

            StringBuilder response = new StringBuilder();
            await foreach (
                var update in answer.WithCancellation(cancellationToken)
            )
            {
                foreach (var content in update.ContentUpdate)
                {
                    if (content.Kind == ChatMessageContentPartKind.Text)
                        response.Append(content.Text);
                }
            }

            return ParseAuditEdits(response.ToString(), maxEdits);
        }

        private static List<AuditEdit> ParseAuditEdits(string response, int maxEdits)
        {
            var result = new List<AuditEdit>();
            if (string.IsNullOrWhiteSpace(response) ||
                response.IndexOf("<none", StringComparison.OrdinalIgnoreCase) >= 0)
                return result;

            MatchCollection matches = Regex.Matches(
                response,
                @"<edit>\s*<find>(.*?)</find>\s*<replace>(.*?)</replace>\s*(?:<reason>(.*?)</reason>\s*)?</edit>",
                RegexOptions.IgnoreCase | RegexOptions.Singleline
            );

            foreach (Match match in matches)
            {
                if (result.Count >= maxEdits)
                    break;

                string find = NormalizeEditText(match.Groups[1].Value);
                string replace = NormalizeEditText(match.Groups[2].Value);
                string reason = match.Groups[3].Success
                    ? NormalizeEditText(match.Groups[3].Value)
                    : string.Empty;

                if (find.Length < 4 || string.Equals(find, replace, StringComparison.Ordinal))
                    continue;

                result.Add(new AuditEdit
                {
                    FindText = find,
                    Replacement = replace,
                    Reason = reason
                });
            }

            return result;
        }

        private int ApplyAuditEdits(
            Word.Document document,
            Word.Range targetRange,
            List<AuditEdit> edits,
            out int skipped,
            out List<AuditEdit> appliedEdits
        )
        {
            skipped = 0;
            appliedEdits = new List<AuditEdit>();

            if (document == null || targetRange == null || edits == null)
            {
                skipped = edits == null ? 0 : edits.Count;
                return 0;
            }

            string currentText = targetRange.Text ?? string.Empty;
            var resolved = new List<ResolvedAuditEdit>();

            foreach (AuditEdit edit in edits)
            {
                if (!IsSafeAuditEdit(edit))
                {
                    skipped++;
                    continue;
                }

                int first = currentText.IndexOf(edit.FindText, StringComparison.Ordinal);
                if (first < 0)
                {
                    skipped++;
                    continue;
                }

                int second = currentText.IndexOf(edit.FindText, first + edit.FindText.Length, StringComparison.Ordinal);
                if (second >= 0)
                {
                    skipped++;
                    continue;
                }

                int editEnd = first + edit.FindText.Length;
                bool overlaps = false;
                foreach (ResolvedAuditEdit existing in resolved)
                {
                    if (existing == null || existing.Edit == null)
                        continue;

                    int existingEnd = existing.RelativeStart + existing.Edit.FindText.Length;
                    if (first < existingEnd && existing.RelativeStart < editEnd)
                    {
                        overlaps = true;
                        break;
                    }
                }

                if (overlaps)
                {
                    skipped++;
                    continue;
                }

                resolved.Add(new ResolvedAuditEdit { Edit = edit, RelativeStart = first });
            }

            if (resolved.Count == 0)
                return 0;

            resolved.Sort((a, b) => b.RelativeStart.CompareTo(a.RelativeStart));
            int originalStart = targetRange.Start;
            int originalEnd = targetRange.End;
            int totalDelta = 0;
            bool originalTrackRevisions = document.TrackRevisions;

            try
            {
                if (!originalTrackRevisions)
                    document.TrackRevisions = true;

                foreach (ResolvedAuditEdit item in resolved)
                {
                    if (item == null || item.Edit == null)
                    {
                        skipped++;
                        continue;
                    }

                    string replacement = item.Edit.Replacement ?? string.Empty;
                    int start = originalStart + item.RelativeStart;
                    int end = start + item.Edit.FindText.Length;
                    Word.Range editRange = document.Range(start, end);
                    if (editRange == null)
                    {
                        skipped++;
                        continue;
                    }

                    editRange.Text = replacement;
                    totalDelta += replacement.Length - item.Edit.FindText.Length;
                    appliedEdits.Add(item.Edit);

                    string appliedLabel = item.Edit.Reason;
                    if (string.IsNullOrWhiteSpace(appliedLabel))
                        appliedLabel = item.Edit.FindText.Length > 80
                            ? item.Edit.FindText.Substring(0, 80) + "…"
                            : item.Edit.FindText;
                    _appliedAuditFixes.Add(appliedLabel);
                }
            }
            finally
            {
                try
                {
                    document.TrackRevisions = originalTrackRevisions;
                }
                catch
                {
                }
            }

            int newEnd = Math.Max(originalStart, originalEnd + totalDelta);
            _auditTargetRange = document.Range(originalStart, newEnd);
            return appliedEdits.Count;
        }

        private static bool IsSafeAuditEdit(AuditEdit edit)
        {
            if (edit == null || string.IsNullOrWhiteSpace(edit.FindText))
                return false;

            string replacement = edit.Replacement ?? string.Empty;
            if (edit.FindText.Length > 1200 || replacement.Length > edit.FindText.Length * 2 + 220)
                return false;

            // A safe editorial change must preserve every number, bracketed citation,
            // DOI and URL from the exact fragment. This second gate protects the document
            // even if the language model ignores the prompt.
            return string.Equals(
                BuildProtectedSignature(edit.FindText),
                BuildProtectedSignature(replacement),
                StringComparison.Ordinal
            );
        }

        private static string BuildProtectedSignature(string text)
        {
            MatchCollection matches = Regex.Matches(
                text ?? string.Empty,
                @"https?://\S+|\bdoi\s*:\s*\S+|\[[^\]\r\n]{1,100}\]|\b\d+(?:[\.,]\d+)?\b",
                RegexOptions.IgnoreCase
            );

            StringBuilder signature = new StringBuilder();
            foreach (Match match in matches)
                signature.AppendLine(match.Value.Trim());
            return signature.ToString();
        }

        private bool TryGetCurrentAuditTarget(out Word.Document document, out Word.Range targetRange)
        {
            document = null;
            targetRange = null;

            if (_auditTargetRange == null || string.IsNullOrWhiteSpace(_lastAuditReport))
            {
                MessageBox.Show(
                    "Сначала запустите «Научный аудит» для выделенного текста и дождитесь готового отчета.",
                    "НеZнайка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return false;
            }

            document = Globals.ThisAddIn.Application.ActiveDocument;
            if (!string.Equals(
                    _auditTargetDocumentIdentity,
                    GetAuditDocumentIdentity(document),
                    StringComparison.OrdinalIgnoreCase
                ))
            {
                MessageBox.Show(
                    "Аудит относится к другому документу. Запустите аудит заново в текущем документе.",
                    "НеZнайка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return false;
            }

            try
            {
                targetRange = _auditTargetRange.Duplicate;
                return targetRange.End > targetRange.Start;
            }
            catch
            {
                MessageBox.Show(
                    "Не удалось найти фрагмент, к которому относился аудит. Запустите аудит заново.",
                    "НеZнайка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                ResetAuditFixState(false);
                return false;
            }
        }

        private void SetAuditControlsBusy(bool busy)
        {
            if (_auditChapterButton != null)
                _auditChapterButton.Enabled = !busy;
            if (_fixAuditButton != null)
                _fixAuditButton.Enabled = !busy && !string.IsNullOrWhiteSpace(_lastAuditReport);
            if (_nextAuditFixButton != null)
                _nextAuditFixButton.Enabled = !busy && !string.IsNullOrWhiteSpace(_lastAuditReport);
            if (_resetAuditButton != null)
                _resetAuditButton.Enabled = !busy && _auditTargetRange != null;
            if (GenerateButton != null)
                GenerateButton.Enabled = !busy;
        }

        private void SetAuditFixButtons(bool enabled)
        {
            if (_fixAuditButton != null)
                _fixAuditButton.Enabled = enabled;
            if (_nextAuditFixButton != null)
                _nextAuditFixButton.Enabled = enabled;
            if (_resetAuditButton != null)
                _resetAuditButton.Enabled = _auditTargetRange != null;
        }

        private void ResetAuditFixState(bool showStatus)
        {
            _auditTargetRange = null;
            _auditTargetDocumentIdentity = null;
            _lastAuditReport = string.Empty;
            _appliedAuditFixes.Clear();
            SetAuditFixButtons(false);

            if (showStatus && _responseLabel != null)
                _responseLabel.Text = "Диалог / ответ — аудит сброшен";
        }

        private void AppendAuditFixNotice(string text)
        {
            if (_responseTextBox == null || _responseTextBox.IsDisposed)
                return;

            if (_responseTextBox.TextLength > 0)
                _responseTextBox.AppendText("\r\n\r\n");
            _responseTextBox.AppendText("[НеZнайка — исправление по аудиту] " + text);
            _responseTextBox.SelectionStart = _responseTextBox.TextLength;
            _responseTextBox.ScrollToCaret();
        }

        private static string NormalizeEditText(string value)
        {
            if (value == null)
                return string.Empty;

            return value
                .Replace("&lt;", "<")
                .Replace("&gt;", ">")
                .Replace("&amp;", "&")
                .Replace("\r\n", "\n")
                .Replace("\r", "\n")
                .Trim();
        }

        private static string GetAuditDocumentIdentity(Word.Document document)
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
