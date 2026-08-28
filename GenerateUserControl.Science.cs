using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
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

        private const string AlwaysUseRagInstruction =
            "Критическое правило для ответа: если TextCraft передал выше RAG-контекст из отмеченных PDF, " +
            "обязательно используй содержащиеся в нем сведения, а не отвечай только по памяти модели. " +
            "При наличии релевантных RAG-фрагментов добавляй конкретные сведения из них и сохраняй ссылки " +
            "на источник и страницу в формате [имя.pdf, с. N], если страница явно указана в контексте. " +
            "Не выдумывай сведения, которых нет в RAG. Если найденные фрагменты не отвечают на вопрос, прямо скажи об этом. " +
            "Если запрос продолжает предыдущий ответ, не повторяй его целиком: дополни новыми релевантными сведениями из RAG.";

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            if (_sciencePanelInitialized)
                return;

            _sciencePanelInitialized = true;
            AddScientificAuditTemplate();
            AddScientificQuickActions();
            AddEvidencePanel();

            // Replace the original chat click handler with the RAG-aware version below.
            // The original handler is still kept in GenerateUserControl.cs for upstream
            // compatibility, but it serializes chat history as one user message and uses
            // a generic follow-up such as "дополни из RAG" as the semantic retrieval query.
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

                List<ChatMessage> messages = BuildRagAwareMessages(
                    userQuery,
                    templateInstruction,
                    localCursorContext
                );

                AppendConversationHeader(userQuery, templateName);
                _responseLabel.Text = "Диалог / ответ — готовлю RAG-контекст…";
                await Task.Yield();

                var streamingAnswer = RAGControl.AskQuestion(
                    new SystemChatMessage(
                        ThisAddIn.SystemPromptLocalization["(GenerateUserControl.cs) _systemPrompt"]
                    ),
                    messages,
                    docRange,
                    GetTemperature()
                );

                _responseLabel.Text = "Диалог / ответ — ожидаю первый токен…";
                string response = await StreamAnswerToPane(streamingAnswer);
                _lastResponseMarkdown = response;
                _lastTemplateName = templateName;

                if (!string.IsNullOrWhiteSpace(response))
                {
                    _conversationTurns.Add(new ConversationTurn(userQuery, response));
                    while (_conversationTurns.Count > MaxConversationTurns)
                        _conversationTurns.RemoveAt(0);

                    _insertButton.Enabled = true;
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
            string localCursorContext
        )
        {
            var messages = new List<ChatMessage>();

            // Preserve real chat roles. Previously the whole conversation, including
            // assistant answers, was wrapped into one UserChatMessage. Small local models
            // tend to echo that block instead of treating the new request as a follow-up.
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

            // This rule is used even for "Свободный запрос". Previously that mode had no
            // instruction forcing the model to consume the RAG context, so a model could
            // answer entirely from its pretrained knowledge despite attached PDFs.
            messages.Add(new UserChatMessage(AlwaysUseRagInstruction));

            if (!string.IsNullOrWhiteSpace(templateInstruction))
                messages.Add(new UserChatMessage(templateInstruction));

            // RAGControl.ProcessInformation() uses messages.Last() as the semantic query.
            // Generic follow-ups like "используй RAG и дополни" carry almost no subject
            // terms, so explicitly include the previous user topic but never the previous
            // assistant answer. This retrieves crystallography chunks for a crystallography
            // follow-up instead of embedding only the words "дополни" and "RAG".
            string retrievalQuery = BuildRagRetrievalQuery(userQuery);
            string finalPrompt =
                "Текущий запрос пользователя: " + userQuery + "\n" +
                "Тема для семантического поиска в RAG: " + retrievalQuery + "\n\n" +
                "Ответь именно на текущий запрос. Если выше передан RAG-контекст, обязательно используй его. " +
                "Для продолжения предыдущего ответа добавляй новые сведения из RAG и не повторяй старый текст дословно.";

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

            // Put the actual subject first. For a request such as "используй материал из
            // RAG и дополни реферат", the useful embedding becomes roughly
            // "расскажи о кристаллографии + используй ...", not just the generic command.
            return previousQuestion + "\n" + current;
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

            // Use the same subject-aware query as the main RAG request, so evidence for a
            // follow-up remains on the previous scientific topic instead of generic words.
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
                    "Проверьте, что нужные источники отмечены галочками в панели «Литература».";
                return;
            }

            StringBuilder text = new StringBuilder();
            int shown = 0;
            foreach (RAGControl.RagEvidenceItem item in evidence)
            {
                if (shown >= 8)
                    break;

                string excerpt = item.Text ?? string.Empty;
                excerpt = excerpt.Replace("\r", " ").Replace("\n", " ").Trim();
                if (excerpt.Length > 420)
                    excerpt = excerpt.Substring(0, 420).TrimEnd() + "…";

                text.AppendLine(item.CitationLabel);
                text.AppendLine(excerpt);
                text.AppendLine();
                shown++;
            }

            _evidenceTextBox.Text = text.ToString().TrimEnd();
        }
    }
}
