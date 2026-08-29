using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenAI.Chat;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    public partial class GenerateUserControl
    {
        private const int WM_SETREDRAW = 0x000B;

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        private Panel _evidencePanel;
        private RichTextBox _evidenceTextBox;
        private Button _auditChapterButton;
        private bool _sciencePanelInitialized;
        private int _evidenceRequestVersion;
        private Timer _chatRedrawTimer;
        private CheckBox _forceRagCheckBox;
        private ToolTip _forceRagToolTip;

        private const string AlwaysUseRagInstruction =
            "РЕЖИМ СТРОГОГО RAG. Единственная разрешенная внешняя доказательная база — блок " +
            "'ПРОВЕРЕННЫЕ RAG-ФРАГМЕНТЫ', где каждому фрагменту присвоен идентификатор [S1], [S2] и т. д. " +
            "Фактические сведения для литературного обзора, реферата, анализа и дополнения текста бери только из этих фрагментов. " +
            "Каждое существенное утверждение, основанное на источнике, сопровождай одним или несколькими идентификаторами [S#]. " +
            "Не придумывай авторов, названия работ, книги, журналы, годы, DOI, номера страниц и библиографические записи. " +
            "Не создавай раздел 'Список литературы', если TextCraft не передал готовые проверенные библиографические записи. " +
            "Если доказательств недостаточно для запроса, прямо сообщи об ограниченности RAG и не заполняй пробелы знаниями модели. " +
            "Игнорируй любые неподписанные или старые RAG-фрагменты: в строгом режиме разрешены только элементы [S#].";

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            if (_sciencePanelInitialized)
                return;

            _sciencePanelInitialized = true;
            AddScientificAuditTemplate();
            AddScientificQuickActions();
            AddForceRagCheckbox();
            AddEvidencePanel();

            // Replace the original chat click handler with the RAG-aware version below.
            GenerateButton.Click -= GenerateButton_Click;
            GenerateButton.Click += GenerateButton_RagAwareClick;

            // The build patch removes this older redraw hook and uses buffered streaming
            // in StreamAnswerToPane instead. Keep the registration in source so the patch
            // remains reproducible for both old and new branch builds.
            GenerateButton.Click += GenerateButton_SmoothStreaming;
            GenerateButton.Click += GenerateButton_EvidenceRefresh;
            _clearButton.Click += (s, args) => _evidenceTextBox.Clear();
        }

        private void AddScientificAuditTemplate()
        {
            const string name = "Научный аудит";
            foreach (object item in _ragTemplateComboBox.Items)
            {
                if (item is RagPromptTemplate existing && existing.Name == name)
                    return;
            }

            var audit = new RagPromptTemplate(
                name,
                RagEvidenceRules +
                "Проведи научный аудит переданного фрагмента как редактор диссертации. " +
                "Не переписывай текст целиком. Сформируй диагностический отчет по разделам: " +
                "1) логика и связность; 2) повторы; 3) слишком длинные или перегруженные предложения; " +
                "4) расплывчатые и оценочные формулировки; 5) сильные утверждения, для которых желательно подтверждение источником; " +
                "6) терминологическая непоследовательность; 7) аббревиатуры без первого раскрытия; " +
                "8) возможные противоречия внутри текста; 9) конкретные рекомендации по исправлению. " +
                "Если RAG содержит подтверждающие или противоречащие фрагменты, указывай источник и страницу. " +
                "Не считать отсутствие фрагмента в RAG доказательством ошибки автора."
            );

            _ragTemplates.Add(audit);
            _ragTemplateComboBox.Items.Add(audit);
        }

        private void AddScientificQuickActions()
        {
            _auditChapterButton = new Button
            {
                Text = "Аудит главы",
                AutoSize = true,
                Height = 28
            };
            _auditChapterButton.Click += AuditChapterButton_Click;
            _quickActionsPanel.Controls.Add(_auditChapterButton);
        }

        private void AddForceRagCheckbox()
        {
            _forceRagCheckBox = new CheckBox
            {
                Text = "RAG: использовать и дополнять текст",
                Checked = true,
                AutoSize = true,
                Height = 26,
                Margin = new Padding(2, 3, 6, 2)
            };

            _forceRagToolTip = new ToolTip();
            _forceRagToolTip.SetToolTip(
                _forceRagCheckBox,
                "Строгий режим: ответ формируется только по найденным фрагментам отмеченных PDF. " +
                "Неподтвержденные ссылки и выдуманный список литературы блокируются."
            );

            _quickActionsPanel.WrapContents = true;
            if (_mainLayout != null && _mainLayout.RowStyles.Count > 2)
                _mainLayout.RowStyles[2].Height = 62F;

            _quickActionsPanel.Controls.Add(_forceRagCheckBox);
            _quickActionsPanel.Controls.SetChildIndex(_forceRagCheckBox, 0);
        }

        private void AddEvidencePanel()
        {
            _evidenceTextBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                DetectUrls = false,
                BackColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Microsoft Sans Serif", 8.5f),
                Margin = new Padding(0)
            };

            Label evidenceLabel = new Label
            {
                Text = "Доказательства из отмеченных PDF:",
                Dock = DockStyle.Top,
                Height = 22,
                TextAlign = ContentAlignment.MiddleLeft
            };

            _evidencePanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 155,
                Padding = new Padding(8, 0, 8, 8)
            };
            _evidencePanel.Controls.Add(_evidenceTextBox);
            _evidencePanel.Controls.Add(evidenceLabel);

            Controls.Add(_evidencePanel);
            Controls.SetChildIndex(_evidencePanel, 0);
        }

        private async void GenerateButton_RagAwareClick(object sender, EventArgs e)
        {
            try
            {
                string userQuery = (PromptTextBox.Text ?? string.Empty).Trim();
                if (userQuery.Length == 0)
                    throw new EmptyTextBoxException(
                        CultureHelper.GetLocalizedString("[GenerateButton_Click] TextBoxEmptyException #1")
                    );

                if (ModelProperties.IsImageModel(ThisAddIn.Model))
                {
                    MessageBox.Show(
                        "Для чата с документом и литературой выберите языковую модель.",
                        "TextCraft",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    return;
                }

                GenerateButton.Enabled = false;
                _insertButton.Enabled = false;
                _copyButton.Enabled = false;

                Word.Range anchorRange = Globals.ThisAddIn.Application.Selection.Range.Duplicate;
                Word.Range docRange = Globals.ThisAddIn.Application.ActiveDocument.Range();
                string localCursorContext = GetLocalCursorContext(anchorRange);
                string templateInstruction = GetSelectedTemplateInstruction();
                string templateName = GetSelectedTemplateName();
                PrepareAuditTargetForRequest(templateName, anchorRange, userQuery);
                bool forceRag = _forceRagCheckBox != null && _forceRagCheckBox.Checked;
                string retrievalQuery = BuildRagRetrievalQuery(userQuery);

                AppendConversationHeader(userQuery, templateName);
                int responseStart = _responseTextBox.TextLength;
                _responseLabel.Text = forceRag
                    ? "Диалог / ответ — проверяю RAG-доказательства…"
                    : "Диалог / ответ — готовлю RAG-контекст…";
                await Task.Yield();

                List<RAGControl.RagEvidenceItem> forcedEvidence = null;
                string groundedEvidenceBlock = string.Empty;

                if (forceRag)
                {
                    RAGControl rag;
                    if (!TryGetRagControlForActiveDocument(out rag))
                    {
                        string noPane =
                            "Строгий RAG остановил генерацию: не удалось связать чат с панелью «Литература» этого документа.";
                        _responseTextBox.AppendText(noPane);
                        _lastResponseMarkdown = noPane;
                        _copyButton.Enabled = true;
                        return;
                    }

                    forcedEvidence = await Task.Run(() => rag.GetRAGEvidence(retrievalQuery, 4));
                    RenderEvidence(forcedEvidence);

                    if (!HasUsableEvidence(forcedEvidence))
                    {
                        string noEvidence =
                            "Строгий RAG остановил генерацию: в отмеченных PDF не найдено подходящих фрагментов по теме запроса. " +
                            "TextCraft не будет дополнять ответ знаниями модели и не будет придумывать литературу. " +
                            "Уточните тему, отметьте другие PDF или отключите строгий RAG.";
                        _responseTextBox.AppendText(noEvidence);
                        _lastResponseMarkdown = noEvidence;
                        _copyButton.Enabled = true;
                        return;
                    }

                    groundedEvidenceBlock = BuildGroundedEvidenceBlock(forcedEvidence);
                }

                List<ChatMessage> messages = BuildRagAwareMessages(
                    userQuery,
                    templateInstruction,
                    localCursorContext,
                    groundedEvidenceBlock,
                    forceRag
                );

                string systemPrompt =
                    ThisAddIn.SystemPromptLocalization["(GenerateUserControl.cs) _systemPrompt"];
                if (forceRag)
                    systemPrompt += "\n\n" + AlwaysUseRagInstruction;

                var streamingAnswer = RAGControl.AskQuestion(
                    new SystemChatMessage(systemPrompt),
                    messages,
                    docRange,
                    GetTemperature()
                );

                _responseLabel.Text = "Диалог / ответ — ожидаю первый токен…";
                string response = await StreamAnswerToPane(streamingAnswer);
                bool responseAccepted = true;

                if (forceRag)
                {
                    response = GroundForcedRagResponse(response, forcedEvidence, out responseAccepted);
                    ReplaceCurrentResponse(responseStart, response);
                }

                _lastResponseMarkdown = response;
                _lastTemplateName = templateName;

                if (!string.IsNullOrWhiteSpace(response))
                {
                    if (responseAccepted)
                    {
                        _conversationTurns.Add(new ConversationTurn(userQuery, response));
                        while (_conversationTurns.Count > MaxConversationTurns)
                            _conversationTurns.RemoveAt(0);

                        _insertButton.Enabled = true;
                    }

                    _copyButton.Enabled = true;
                }

                PromptTextBox.Clear();
                PromptTextBox.Focus();
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
            finally
            {
                GenerateButton.Enabled = true;
                if (_responseLabel != null)
                    _responseLabel.Text = "Диалог / ответ:";
            }
        }

        private List<ChatMessage> BuildRagAwareMessages(
            string userQuery,
            string templateInstruction,
            string localCursorContext,
            string groundedEvidenceBlock,
            bool forceRag
        )
        {
            var messages = new List<ChatMessage>();

            if (_conversationTurns.Count > 0)
            {
                int answerTokens = Math.Max(256, (int)(ThisAddIn.ContextLength * 0.07));
                int questionTokens = Math.Max(96, (int)(ThisAddIn.ContextLength * 0.02));

                foreach (ConversationTurn turn in _conversationTurns)
                {
                    messages.Add(
                        new UserChatMessage(
                            CommonUtils.SubstringTokens(turn.Question ?? string.Empty, questionTokens)
                        )
                    );
                    messages.Add(
                        new AssistantChatMessage(
                            CommonUtils.SubstringTokens(turn.Answer ?? string.Empty, answerTokens)
                        )
                    );
                }
            }

            if (!string.IsNullOrWhiteSpace(localCursorContext))
            {
                messages.Add(
                    new UserChatMessage(
                        "Локальный контекст вокруг курсора из текущего Word-документа. " +
                        "Используй его для связи ответа с текущим разделом, но не считай его внешним доказательством:\n\n\"" +
                        localCursorContext +
                        "\""
                    )
                );
            }

            if (forceRag)
            {
                messages.Add(new UserChatMessage(AlwaysUseRagInstruction));
                messages.Add(new UserChatMessage(groundedEvidenceBlock));
            }

            if (!string.IsNullOrWhiteSpace(templateInstruction))
                messages.Add(new UserChatMessage(templateInstruction));

            string retrievalQuery = BuildRagRetrievalQuery(userQuery);
            string ragBehavior = forceRag
                ? "Строгий RAG включен. Пиши только по проверенным фрагментам [S#]. " +
                  "Не создавай библиографию и не используй сведения, которых нет в этих фрагментах. " +
                  "Для каждой опоры на источник ставь [S#]."
                : "Ответь на текущий запрос с учетом предыдущего диалога. RAG-контекст можно использовать как дополнительный источник, если он релевантен.";

            string finalPrompt =
                "Текущий запрос пользователя: " + userQuery + "\n" +
                "Тема для семантического поиска в RAG: " + retrievalQuery + "\n\n" +
                ragBehavior;

            messages.Add(new UserChatMessage(finalPrompt));
            return messages;
        }

        private string BuildRagRetrievalQuery(string userQuery)
        {
            string current = (userQuery ?? string.Empty).Trim();
            if (_conversationTurns.Count == 0)
                return current;

            ConversationTurn previous = _conversationTurns[_conversationTurns.Count - 1];
            string previousQuestion = (previous.Question ?? string.Empty).Trim();
            if (previousQuestion.Length == 0)
                return current;

            return previousQuestion + "\n" + current;
        }

        private static bool HasUsableEvidence(List<RAGControl.RagEvidenceItem> evidence)
        {
            if (evidence == null || evidence.Count == 0)
                return false;

            foreach (RAGControl.RagEvidenceItem item in evidence)
            {
                if (item != null && !string.IsNullOrWhiteSpace(item.Text))
                    return true;
            }

            return false;
        }

        private static string BuildGroundedEvidenceBlock(List<RAGControl.RagEvidenceItem> evidence)
        {
            StringBuilder block = new StringBuilder();
            block.AppendLine("ПРОВЕРЕННЫЕ RAG-ФРАГМЕНТЫ — ЕДИНСТВЕННАЯ РАЗРЕШЕННАЯ ВНЕШНЯЯ ДОКАЗАТЕЛЬНАЯ БАЗА:");
            block.AppendLine("Цитируй только идентификаторами [S1], [S2] и т. д. Не придумывай библиографические данные.");
            block.AppendLine();

            int shown = 0;
            foreach (RAGControl.RagEvidenceItem item in evidence)
            {
                if (item == null || string.IsNullOrWhiteSpace(item.Text))
                    continue;
                if (shown >= 10)
                    break;

                shown++;
                string excerpt = item.Text.Replace("\r", " ").Replace("\n", " ").Trim();
                if (excerpt.Length > 900)
                    excerpt = excerpt.Substring(0, 900).TrimEnd() + "…";

                block.Append("[S").Append(shown).Append("] ");
                block.Append(item.CitationLabel).AppendLine();
                block.AppendLine(excerpt);
                block.AppendLine();
            }

            return block.ToString();
        }

        private static string GroundForcedRagResponse(
            string response,
            List<RAGControl.RagEvidenceItem> evidence,
            out bool accepted
        )
        {
            accepted = false;
            string grounded = (response ?? string.Empty).Trim();
            if (grounded.Length == 0)
                return grounded;

            int maxSources = 0;
            if (evidence != null)
            {
                foreach (RAGControl.RagEvidenceItem item in evidence)
                {
                    if (item != null && !string.IsNullOrWhiteSpace(item.Text))
                        maxSources++;
                    if (maxSources >= 10)
                        break;
                }
            }

            bool usedVerifiedMarker = false;
            for (int i = 1; i <= maxSources; i++)
            {
                string marker = "[S" + i + "]";
                if (grounded.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0)
                    usedVerifiedMarker = true;
            }

            if (!usedVerifiedMarker)
            {
                return
                    "⚠ Строгий RAG отклонил ответ модели: в тексте нет ни одной ссылки на проверенный фрагмент [S#]. " +
                    "Ответ не разрешен для вставки, потому что модель могла использовать собственные знания вместо отмеченных PDF. " +
                    "Уточните запрос или отключите строгий RAG.";
            }

            int evidenceIndex = 0;
            if (evidence != null)
            {
                foreach (RAGControl.RagEvidenceItem item in evidence)
                {
                    if (item == null || string.IsNullOrWhiteSpace(item.Text))
                        continue;
                    if (evidenceIndex >= 10)
                        break;

                    evidenceIndex++;
                    grounded = Regex.Replace(
                        grounded,
                        @"\[S" + evidenceIndex + @"\]",
                        item.CitationLabel.Replace("$", "$$"),
                        RegexOptions.IgnoreCase
                    );
                }
            }

            // Any S-marker left at this point was not part of the verified evidence set.
            grounded = Regex.Replace(
                grounded,
                @"\[S\d+\]",
                "[неподтвержденная RAG-ссылка удалена]",
                RegexOptions.IgnoreCase
            );

            // Numeric pseudo-citations such as [1] or [2, с. 21] are not allowed in strict
            // mode because the model can fabricate the corresponding bibliography.
            grounded = Regex.Replace(
                grounded,
                @"\[\d+\s*(?:,\s*с\.\s*\d+)?\]",
                "[неподтвержденная ссылка удалена]",
                RegexOptions.IgnoreCase
            );

            int bibliographyIndex = grounded.IndexOf("Список литературы", StringComparison.OrdinalIgnoreCase);
            if (bibliographyIndex >= 0)
            {
                grounded = grounded.Substring(0, bibliographyIndex).TrimEnd() +
                    "\n\n[TextCraft: список литературы не добавлен, поскольку его библиографические записи не были подтверждены RAG.]";
            }

            accepted = true;
            return grounded.Trim();
        }

        private void ReplaceCurrentResponse(int responseStart, string response)
        {
            if (_responseTextBox == null || _responseTextBox.IsDisposed)
                return;

            int safeStart = Math.Max(0, Math.Min(responseStart, _responseTextBox.TextLength));
            int length = _responseTextBox.TextLength - safeStart;
            _responseTextBox.Select(safeStart, length);
            _responseTextBox.SelectedText = response ?? string.Empty;
            _responseTextBox.SelectionStart = _responseTextBox.TextLength;
            _responseTextBox.SelectionLength = 0;
            _responseTextBox.ScrollToCaret();
        }

        private async void GenerateButton_SmoothStreaming(object sender, EventArgs e)
        {
            await Task.Yield();
            if (GenerateButton.Enabled || _responseTextBox == null || !_responseTextBox.IsHandleCreated)
                return;

            StopChatRedrawTimer();
            SetChatRedraw(false);

            _chatRedrawTimer = new Timer { Interval = 220 };
            _chatRedrawTimer.Tick += (s, args) =>
            {
                if (GenerateButton.Enabled)
                {
                    StopChatRedrawTimer();
                    SetChatRedraw(true);
                    _responseTextBox.Refresh();
                    return;
                }

                SetChatRedraw(true);
                _responseTextBox.Refresh();
                SetChatRedraw(false);
            };
            _chatRedrawTimer.Start();
        }

        private void StopChatRedrawTimer()
        {
            if (_chatRedrawTimer == null)
                return;

            _chatRedrawTimer.Stop();
            _chatRedrawTimer.Dispose();
            _chatRedrawTimer = null;
        }

        private void SetChatRedraw(bool enabled)
        {
            if (_responseTextBox == null || !_responseTextBox.IsHandleCreated)
                return;

            SendMessage(
                _responseTextBox.Handle,
                WM_SETREDRAW,
                enabled ? new IntPtr(1) : IntPtr.Zero,
                IntPtr.Zero
            );
        }

        private void AuditChapterButton_Click(object sender, EventArgs e)
        {
            try
            {
                Word.Selection selection = Globals.ThisAddIn.Application.Selection;
                if (selection == null || selection.End <= selection.Start)
                {
                    MessageBox.Show(
                        "Выделите главу, раздел или несколько абзацев для научного аудита.",
                        "TextCraft",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    return;
                }

                SelectTemplate("Научный аудит");

                int maxTokens = Math.Max(900, (int)(ThisAddIn.ContextLength * 0.42));
                PromptTextBox.Text = CommonUtils.SubstringTokens(
                    (selection.Text ?? string.Empty).Trim(),
                    maxTokens
                );
                GenerateButton.PerformClick();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private async void GenerateButton_EvidenceRefresh(object sender, EventArgs e)
        {
            string query = (PromptTextBox.Text ?? string.Empty).Trim();
            if (query.Length == 0)
                return;

            string evidenceQuery = BuildRagRetrievalQuery(query);

            int requestVersion = ++_evidenceRequestVersion;
            _evidenceTextBox.Text = "Поиск подтверждающих фрагментов…";

            for (int i = 0; i < 240 && !GenerateButton.Enabled; i++)
                await Task.Delay(250);

            if (requestVersion != _evidenceRequestVersion)
                return;

            try
            {
                RAGControl rag;
                if (!TryGetRagControlForActiveDocument(out rag))
                {
                    _evidenceTextBox.Text = "Не удалось связать чат с панелью литературы этого документа.";
                    return;
                }

                List<RAGControl.RagEvidenceItem> evidence = await Task.Run(
                    () => rag.GetRAGEvidence(evidenceQuery, 2)
                );

                if (requestVersion != _evidenceRequestVersion)
                    return;

                RenderEvidence(evidence);
            }
            catch (Exception ex)
            {
                _evidenceTextBox.Text = "Не удалось получить фрагменты доказательств: " + ex.Message;
            }
        }

        private static bool TryGetRagControlForActiveDocument(out RAGControl rag)
        {
            rag = null;
            Word.Document active = Globals.ThisAddIn.Application.ActiveDocument;
            if (active == null)
                return false;

            if (ThisAddIn.AllTaskPanes.TryGetValue(active, out var direct))
            {
                rag = direct.Item3;
                return rag != null;
            }

            string activeFullName = string.Empty;
            string activeName = string.Empty;
            try { activeFullName = active.FullName ?? string.Empty; } catch { }
            try { activeName = active.Name ?? string.Empty; } catch { }

            foreach (var entry in ThisAddIn.AllTaskPanes)
            {
                try
                {
                    string candidateFullName = entry.Key.FullName ?? string.Empty;
                    string candidateName = entry.Key.Name ?? string.Empty;
                    if ((!string.IsNullOrEmpty(activeFullName) && string.Equals(candidateFullName, activeFullName, StringComparison.OrdinalIgnoreCase)) ||
                        (!string.IsNullOrEmpty(activeName) && string.Equals(candidateName, activeName, StringComparison.OrdinalIgnoreCase)))
                    {
                        rag = entry.Value.Item3;
                        return rag != null;
                    }
                }
                catch
                {
                    // Closed/stale COM document entry; continue with the next pane.
                }
            }

            if (ThisAddIn.AllTaskPanes.Count == 1)
            {
                foreach (var entry in ThisAddIn.AllTaskPanes)
                {
                    rag = entry.Value.Item3;
                    return rag != null;
                }
            }

            return false;
        }

        private void RenderEvidence(List<RAGControl.RagEvidenceItem> evidence)
        {
            if (evidence == null || evidence.Count == 0)
            {
                _evidenceTextBox.Text =
                    "Подходящие фрагменты в отмеченных PDF не найдены. " +
                    "В строгом RAG генерация будет остановлена, чтобы модель не выдумывала источники.";
                return;
            }

            StringBuilder text = new StringBuilder();
            int shown = 0;
            foreach (RAGControl.RagEvidenceItem item in evidence)
            {
                if (item == null || string.IsNullOrWhiteSpace(item.Text))
                    continue;
                if (shown >= 8)
                    break;

                string excerpt = item.Text.Replace("\r", " ").Replace("\n", " ").Trim();
                if (excerpt.Length > 420)
                    excerpt = excerpt.Substring(0, 420).TrimEnd() + "…";

                text.Append("[S").Append(shown + 1).Append("] ");
                text.AppendLine(item.CitationLabel);
                text.AppendLine(excerpt);
                text.AppendLine();
                shown++;
            }

            if (shown == 0)
            {
                _evidenceTextBox.Text =
                    "Подходящие фрагменты в отмеченных PDF не найдены. " +
                    "В строгом RAG генерация будет остановлена, чтобы модель не выдумывала источники.";
                return;
            }

            _evidenceTextBox.Text = text.ToString().TrimEnd();
        }
    }
}
