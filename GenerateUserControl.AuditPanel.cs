using System;
using System.ClientModel;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
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
        private Panel _auditReviewPanel;
        private TableLayoutPanel _auditReviewLayout;
        private Label _auditReviewHeader;
        private CheckedListBox _auditIssueList;
        private RichTextBox _auditReasonTextBox;
        private RichTextBox _auditBeforeTextBox;
        private RichTextBox _auditAfterTextBox;
        private FlowLayoutPanel _auditReviewActions;
        private Button _auditGoToButton;
        private Button _auditApplySelectedButton;
        private Button _auditSelectSafeButton;
        private Button _auditClosePanelButton;
        private readonly List<AuditReviewIssue> _auditReviewIssues = new List<AuditReviewIssue>();
        private bool _auditReviewBusy;

        private sealed class AuditReviewIssue
        {
            public string Category { get; set; }
            public string FindText { get; set; }
            public string Replacement { get; set; }
            public string Reason { get; set; }
            public bool ModelMarkedSafe { get; set; }
            public bool AutoApplicable { get; set; }
            public bool Applied { get; set; }

            public override string ToString()
            {
                string state = Applied ? "[применено] " : (AutoApplicable ? "[авто] " : "[ручная] ");
                string category = string.IsNullOrWhiteSpace(Category) ? "Замечание" : Category.Trim();
                string reason = string.IsNullOrWhiteSpace(Reason) ? FindText : Reason.Trim();
                reason = Regex.Replace(reason ?? string.Empty, @"\s+", " ");
                if (reason.Length > 105)
                    reason = reason.Substring(0, 105).TrimEnd() + "…";
                return state + "[" + category + "] " + reason;
            }
        }

        private sealed class AuditReviewApplyItem
        {
            public AuditReviewIssue Issue { get; set; }
            public int Position { get; set; }
        }

        private void InitializeAuditReviewPanel()
        {
            if (_auditReviewPanel != null)
                return;

            _auditReviewHeader = new Label
            {
                Text = "Замечания аудита",
                Dock = DockStyle.Fill,
                Height = 24,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold),
                ForeColor = SystemColors.ControlText,
                BackColor = SystemColors.Control
            };

            _auditIssueList = new CheckedListBox
            {
                Dock = DockStyle.Fill,
                CheckOnClick = true,
                IntegralHeight = false,
                HorizontalScrollbar = true
            };
            _auditIssueList.SelectedIndexChanged += AuditIssueList_SelectedIndexChanged;
            _auditIssueList.ItemCheck += AuditIssueList_ItemCheck;

            _auditReasonTextBox = CreateAuditPreviewBox();
            _auditBeforeTextBox = CreateAuditPreviewBox();
            _auditAfterTextBox = CreateAuditPreviewBox();

            _auditGoToButton = new Button
            {
                Text = "К месту",
                AutoSize = true,
                Height = 28,
                Enabled = false
            };
            _auditGoToButton.Click += AuditGoToButton_Click;

            _auditApplySelectedButton = new Button
            {
                Text = "Применить отмеченные",
                AutoSize = true,
                Height = 28,
                Enabled = false
            };
            _auditApplySelectedButton.Click += AuditApplySelectedButton_Click;

            _auditSelectSafeButton = new Button
            {
                Text = "Безопасные",
                AutoSize = true,
                Height = 28,
                Enabled = false
            };
            _auditSelectSafeButton.Click += AuditSelectSafeButton_Click;

            _auditClosePanelButton = new Button
            {
                Text = "Закрыть",
                AutoSize = true,
                Height = 28
            };
            _auditClosePanelButton.Click += (s, e) => HideAuditReviewPanel();

            _auditReviewActions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = false,
                Height = 62,
                WrapContents = true,
                FlowDirection = FlowDirection.LeftToRight,
                Margin = new Padding(0)
            };
            _auditReviewActions.Controls.Add(_auditGoToButton);
            _auditReviewActions.Controls.Add(_auditApplySelectedButton);
            _auditReviewActions.Controls.Add(_auditSelectSafeButton);
            _auditReviewActions.Controls.Add(_auditClosePanelButton);

            _auditReviewLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 9,
                Padding = new Padding(0)
            };
            _auditReviewLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            _auditReviewLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 25F));
            _auditReviewLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 94F));
            _auditReviewLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 19F));
            _auditReviewLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46F));
            _auditReviewLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 19F));
            _auditReviewLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 52F));
            _auditReviewLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 19F));
            _auditReviewLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 52F));
            _auditReviewLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            _auditReviewLayout.Controls.Add(_auditReviewHeader, 0, 0);
            _auditReviewLayout.Controls.Add(_auditIssueList, 0, 1);
            _auditReviewLayout.Controls.Add(CreateAuditSectionLabel("Почему:"), 0, 2);
            _auditReviewLayout.Controls.Add(_auditReasonTextBox, 0, 3);
            _auditReviewLayout.Controls.Add(CreateAuditSectionLabel("Было:"), 0, 4);
            _auditReviewLayout.Controls.Add(_auditBeforeTextBox, 0, 5);
            _auditReviewLayout.Controls.Add(CreateAuditSectionLabel("Стало:"), 0, 6);
            _auditReviewLayout.Controls.Add(_auditAfterTextBox, 0, 7);
            _auditReviewLayout.Controls.Add(_auditReviewActions, 0, 8);

            _auditReviewPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 390,
                Padding = new Padding(8, 6, 8, 6),
                Visible = false,
                BorderStyle = BorderStyle.FixedSingle
            };
            _auditReviewPanel.Controls.Add(_auditReviewLayout);

            Controls.Add(_auditReviewPanel);
            Controls.SetChildIndex(_auditReviewPanel, 0);

            if (_clearButton != null)
                _clearButton.Click += (s, e) => ResetAuditReviewPanelState(false);
        }

        private static RichTextBox CreateAuditPreviewBox()
        {
            return new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                DetectUrls = false,
                BackColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Microsoft Sans Serif", 8.5F),
                Margin = new Padding(0)
            };
        }

        private static Label CreateAuditSectionLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                Height = 19,
                TextAlign = ContentAlignment.BottomLeft,
                Margin = new Padding(0),
                ForeColor = SystemColors.ControlText,
                BackColor = SystemColors.Control
            };
        }

        private void PrepareAuditReviewForNewRun()
        {
            _auditReviewIssues.Clear();
            if (_auditIssueList != null)
                _auditIssueList.Items.Clear();

            ClearAuditPreview();
            if (_auditReviewHeader != null)
                _auditReviewHeader.Text = "Замечания аудита — ожидаю результат…";

            if (_auditReviewPanel != null)
                _auditReviewPanel.Visible = false;
            if (_evidencePanel != null)
                _evidencePanel.Visible = true;
        }

        private bool HasPendingAuditReview()
        {
            return _auditTargetRange != null && !string.IsNullOrWhiteSpace(_lastAuditReport);
        }

        private List<AuditEdit> GetPendingSafeAuditReviewEdits(int maxEdits)
        {
            var result = new List<AuditEdit>();
            if (maxEdits <= 0)
                return result;

            foreach (AuditReviewIssue issue in _auditReviewIssues)
            {
                if (issue == null || issue.Applied || !issue.AutoApplicable)
                    continue;

                result.Add(new AuditEdit
                {
                    FindText = issue.FindText ?? string.Empty,
                    Replacement = issue.Replacement ?? string.Empty,
                    Reason = issue.Reason ?? string.Empty
                });

                if (result.Count >= maxEdits)
                    break;
            }

            return result;
        }

        private void MarkAuditReviewEditsApplied(IEnumerable<AuditEdit> edits)
        {
            if (edits == null)
                return;

            var appliedKeys = new HashSet<string>(StringComparer.Ordinal);
            foreach (AuditEdit edit in edits)
            {
                if (edit != null)
                    appliedKeys.Add((edit.FindText ?? string.Empty) + "\n" + (edit.Replacement ?? string.Empty));
            }

            foreach (AuditReviewIssue issue in _auditReviewIssues)
            {
                if (issue == null)
                    continue;

                string key = (issue.FindText ?? string.Empty) + "\n" + (issue.Replacement ?? string.Empty);
                if (appliedKeys.Contains(key))
                    issue.Applied = true;
            }

            RenderAuditReviewIssues(true);
            int safeCount = _auditReviewIssues.Count(i => i.AutoApplicable && !i.Applied);
            if (_auditReviewHeader != null)
            {
                _auditReviewHeader.Text =
                    "Замечания аудита: " + _auditReviewIssues.Count +
                    " (безопасных осталось: " + safeCount + ")";
            }
        }

        private async Task BuildAuditReviewPanelAsync()
        {
            if (_auditReviewBusy || !HasPendingAuditReview())
                return;

            Word.Document document;
            Word.Range targetRange;
            if (!TryGetCurrentAuditTarget(out document, out targetRange))
                return;

            string currentText = targetRange.Text ?? string.Empty;
            if (string.IsNullOrWhiteSpace(currentText))
                return;

            try
            {
                _auditReviewBusy = true;
                SetAuditFixButtons(false);
                ShowAuditReviewPanel();
                ClearAuditPreview();
                _auditReviewHeader.Text = "Замечания аудита — разбираю отчет…";
                _auditReasonTextBox.Text = "Подождите: НеZнайка преобразует отчет в отдельные замечания…";
                if (_responseLabel != null)
                    _responseLabel.Text = "Аудит — структурирую замечания…";
                SetAuditReviewControlsEnabled(false);
                _auditReviewPanel.Refresh();
                await Task.Yield();

                List<AuditReviewIssue> issues = await GenerateAuditReviewIssuesAsync(
                    currentText,
                    _lastAuditReport,
                    20
                );

                _auditReviewIssues.Clear();
                _auditReviewIssues.AddRange(issues);
                RenderAuditReviewIssues(true);

                int safeCount = _auditReviewIssues.Count(i => i.AutoApplicable && !i.Applied);
                _auditReviewHeader.Text =
                    "Замечания аудита: " + _auditReviewIssues.Count +
                    " (безопасных: " + safeCount + ")";

                if (_auditReviewIssues.Count == 0)
                {
                    _auditReasonTextBox.Text =
                        "Структурированные замечания не получены. Диагностический отчет остается в окне ответа.";
                }
            }
            catch (OperationCanceledException)
            {
                _auditReviewHeader.Text = "Замечания аудита — разбор остановлен";
            }
            catch (Exception ex)
            {
                _auditReviewHeader.Text = "Замечания аудита — ошибка разбора";
                _auditReasonTextBox.Text = ex.Message;
            }
            finally
            {
                _auditReviewBusy = false;
                SetAuditReviewControlsEnabled(true);
                SetAuditFixButtons(HasPendingAuditReview());
            }
        }

        private async Task<List<AuditReviewIssue>> GenerateAuditReviewIssuesAsync(
            string currentText,
            string auditReport,
            int maxIssues
        )
        {
            int textTokens = Math.Max(1200, (int)(ThisAddIn.ContextLength * 0.38));
            int auditTokens = Math.Max(800, (int)(ThisAddIn.ContextLength * 0.24));
            string boundedText = CommonUtils.SubstringTokens(currentText, textTokens);
            string boundedAudit = CommonUtils.SubstringTokens(auditReport ?? string.Empty, auditTokens);

            string systemPrompt =
                "Ты научный редактор диссертации. Преобразуй готовый диагностический аудит в структурированный список замечаний. " +
                "Разрешенные категории: Логика, Повтор, Стиль, Терминология, Язык, Источник, Противоречие, Аббревиатура. " +
                "Для каждого замечания FIND должен быть дословной уникальной подстрокой переданного текста, по возможности одним предложением. " +
                "SAFE=true разрешено только для локальной редакторской правки: повтор, громоздкий синтаксис, неясная связка, расплывчатая формулировка, " +
                "очевидная языковая ошибка или терминологическая непоследовательность, если смысл и факты не меняются. " +
                "SAFE=false обязательно для замечаний о необходимости источника, недостатке доказательств, фактическом противоречии, числах, формулах, цитатах, " +
                "именах, датах, DOI, URL, библиографических ссылках, научных выводах и для любой правки, требующей решения автора. " +
                "При SAFE=false оставь REPLACE пустым. Не добавляй новых фактов. Не используй Markdown и не пиши ничего вне тегов.";

            string userPrompt =
                "Аудит:\n" + boundedAudit + "\n\n" +
                "Текст:\n<<<TEXT>>>\n" + boundedText + "\n<<<END TEXT>>>\n\n" +
                "Верни не более " + maxIssues + " элементов строго в формате:\n" +
                "<issue>\n" +
                "<type>категория</type>\n" +
                "<find>дословный фрагмент текста</find>\n" +
                "<replace>безопасная замена или пусто</replace>\n" +
                "<reason>краткое объяснение замечания</reason>\n" +
                "<safe>true или false</safe>\n" +
                "</issue>\n" +
                "Если замечание нельзя привязать к конкретному фрагменту текста, не включай его.";

            var cancellationToken = GetAuditOperationToken();

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
                new ChatCompletionOptions { Temperature = 0.03f },
                cancellationToken
            );
            if (answer == null)
                throw new InvalidOperationException("Модель не вернула поток структурированных замечаний.");

            StringBuilder response = new StringBuilder();
            await foreach (
                var update in answer.WithCancellation(cancellationToken)
            )
            {
                foreach (var part in update.ContentUpdate)
                {
                    if (part.Kind == ChatMessageContentPartKind.Text)
                        response.Append(part.Text);
                }
            }

            return ParseAuditReviewIssues(response.ToString(), currentText, maxIssues);
        }

        private static List<AuditReviewIssue> ParseAuditReviewIssues(
            string response,
            string currentText,
            int maxIssues
        )
        {
            var result = new List<AuditReviewIssue>();
            if (string.IsNullOrWhiteSpace(response))
                return result;

            MatchCollection matches = Regex.Matches(
                response,
                @"<issue>\s*<type>(.*?)</type>\s*<find>(.*?)</find>\s*<replace>(.*?)</replace>\s*<reason>(.*?)</reason>\s*<safe>(.*?)</safe>\s*</issue>",
                RegexOptions.IgnoreCase | RegexOptions.Singleline
            );

            foreach (Match match in matches)
            {
                if (result.Count >= maxIssues)
                    break;

                string category = NormalizeEditText(match.Groups[1].Value);
                string find = NormalizeEditText(match.Groups[2].Value);
                string replacement = NormalizeEditText(match.Groups[3].Value);
                string reason = NormalizeEditText(match.Groups[4].Value);
                string safeText = NormalizeEditText(match.Groups[5].Value);

                if (find.Length < 4)
                    continue;

                int first = currentText.IndexOf(find, StringComparison.Ordinal);
                if (first < 0)
                    continue;

                int second = currentText.IndexOf(find, first + find.Length, StringComparison.Ordinal);
                bool unique = second < 0;
                bool modelSafe =
                    safeText.Equals("true", StringComparison.OrdinalIgnoreCase) ||
                    safeText.Equals("yes", StringComparison.OrdinalIgnoreCase) ||
                    safeText.Equals("да", StringComparison.OrdinalIgnoreCase) ||
                    safeText == "1";

                var issue = new AuditReviewIssue
                {
                    Category = category,
                    FindText = find,
                    Replacement = replacement,
                    Reason = reason,
                    ModelMarkedSafe = modelSafe,
                    Applied = false
                };

                if (modelSafe && unique && !string.IsNullOrWhiteSpace(replacement))
                {
                    issue.AutoApplicable = IsSafeAuditEdit(
                        new AuditEdit
                        {
                            FindText = find,
                            Replacement = replacement,
                            Reason = reason
                        }
                    );
                }

                result.Add(issue);
            }

            return result;
        }

        private void RenderAuditReviewIssues(bool selectFirst)
        {
            if (_auditIssueList == null)
                return;

            _auditIssueList.BeginUpdate();
            try
            {
                _auditIssueList.Items.Clear();
                for (int i = 0; i < _auditReviewIssues.Count; i++)
                {
                    AuditReviewIssue issue = _auditReviewIssues[i];
                    int index = _auditIssueList.Items.Add(issue);
                    if (issue.AutoApplicable && !issue.Applied)
                        _auditIssueList.SetItemChecked(index, true);
                }
            }
            finally
            {
                _auditIssueList.EndUpdate();
            }

            if (selectFirst && _auditIssueList.Items.Count > 0)
                _auditIssueList.SelectedIndex = 0;
            else if (_auditIssueList.Items.Count == 0)
                ClearAuditPreview();

            UpdateAuditActionButtons();
        }

        private void AuditIssueList_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _auditIssueList.Items.Count)
                return;

            AuditReviewIssue issue = _auditIssueList.Items[e.Index] as AuditReviewIssue;
            if (issue != null && (!issue.AutoApplicable || issue.Applied) && e.NewValue != CheckState.Unchecked)
            {
                e.NewValue = CheckState.Unchecked;
                BeginInvoke(new Action(() =>
                {
                    _auditReviewHeader.Text =
                        "Это замечание требует решения автора и не будет применено автоматически.";
                }));
            }

            BeginInvoke(new Action(UpdateAuditActionButtons));
        }

        private void AuditIssueList_SelectedIndexChanged(object sender, EventArgs e)
        {
            AuditReviewIssue issue = _auditIssueList.SelectedItem as AuditReviewIssue;
            ShowAuditIssuePreview(issue);
        }

        private void ShowAuditIssuePreview(AuditReviewIssue issue)
        {
            if (issue == null)
            {
                ClearAuditPreview();
                return;
            }

            _auditReasonTextBox.Text = issue.Reason ?? string.Empty;
            _auditBeforeTextBox.Text = issue.FindText ?? string.Empty;

            if (issue.Applied)
                _auditAfterTextBox.Text = "Правка уже применена через рецензирование Word.";
            else if (issue.AutoApplicable)
                _auditAfterTextBox.Text = issue.Replacement ?? string.Empty;
            else
                _auditAfterTextBox.Text =
                    "Автоматическая замена запрещена. Требуется решение автора или проверка источников.";

            _auditGoToButton.Enabled = !string.IsNullOrWhiteSpace(issue.FindText);
            UpdateAuditActionButtons();
        }

        private void ClearAuditPreview()
        {
            if (_auditReasonTextBox != null)
                _auditReasonTextBox.Clear();
            if (_auditBeforeTextBox != null)
                _auditBeforeTextBox.Clear();
            if (_auditAfterTextBox != null)
                _auditAfterTextBox.Clear();
            if (_auditGoToButton != null)
                _auditGoToButton.Enabled = false;
            UpdateAuditActionButtons();
        }

        private void AuditGoToButton_Click(object sender, EventArgs e)
        {
            AuditReviewIssue issue = _auditIssueList.SelectedItem as AuditReviewIssue;
            if (issue == null)
                return;

            Word.Range range;
            if (!TryResolveAuditIssueRange(issue, out range))
            {
                MessageBox.Show(
                    "Фрагмент больше не найден однозначно. Возможно, текст уже был изменен. Выполните аудит заново.",
                    "НеZнайка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return;
            }

            try
            {
                range.Select();
                Word.Window window = Globals.ThisAddIn.Application.ActiveWindow;
                if (window != null)
                    window.ScrollIntoView(range, true);
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private async void AuditApplySelectedButton_Click(object sender, EventArgs e)
        {
            if (_auditReviewBusy)
                return;

            Word.Document document;
            Word.Range targetRange;
            if (!TryGetCurrentAuditTarget(out document, out targetRange))
                return;

            var selected = new List<AuditReviewApplyItem>();
            string currentText = targetRange.Text ?? string.Empty;

            foreach (object checkedItem in _auditIssueList.CheckedItems)
            {
                AuditReviewIssue issue = checkedItem as AuditReviewIssue;
                if (issue == null || issue.Applied || !issue.AutoApplicable)
                    continue;

                int first = currentText.IndexOf(issue.FindText, StringComparison.Ordinal);
                if (first < 0)
                    continue;
                int second = currentText.IndexOf(
                    issue.FindText,
                    first + issue.FindText.Length,
                    StringComparison.Ordinal
                );
                if (second >= 0)
                    continue;

                selected.Add(new AuditReviewApplyItem { Issue = issue, Position = first });
            }

            if (selected.Count == 0)
            {
                MessageBox.Show(
                    "Отметьте одну или несколько безопасных правок.",
                    "НеZнайка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return;
            }

            selected.Sort((a, b) => b.Position.CompareTo(a.Position));

            try
            {
                _auditReviewBusy = true;
                SetAuditReviewControlsEnabled(false);
                int applied = 0;
                int skipped = 0;

                foreach (AuditReviewApplyItem item in selected)
                {
                    int localSkipped;
                    int count = ApplyAuditEdits(
                        document,
                        _auditTargetRange.Duplicate,
                        new List<AuditEdit>
                        {
                            new AuditEdit
                            {
                                FindText = item.Issue.FindText,
                                Replacement = item.Issue.Replacement,
                                Reason = item.Issue.Reason
                            }
                        },
                        out localSkipped,
                        out List<AuditEdit> localAppliedEdits
                    );

                    if (count == 1)
                    {
                        item.Issue.Applied = true;
                        applied++;
                    }
                    else
                    {
                        skipped++;
                    }
                }

                RenderAuditReviewIssues(false);
                _auditReviewHeader.Text =
                    "Замечания аудита — применено: " + applied +
                    (skipped > 0 ? ", пропущено: " + skipped : string.Empty);
                AppendAuditFixNotice(
                    "Из панели замечаний применено через рецензирование Word: " + applied +
                    (skipped > 0 ? ". Пропущено: " + skipped + "." : ".")
                );
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
            finally
            {
                _auditReviewBusy = false;
                SetAuditReviewControlsEnabled(true);
            }

            await Task.Yield();
        }

        private void AuditSelectSafeButton_Click(object sender, EventArgs e)
        {
            if (_auditIssueList == null)
                return;

            for (int i = 0; i < _auditIssueList.Items.Count; i++)
            {
                AuditReviewIssue issue = _auditIssueList.Items[i] as AuditReviewIssue;
                _auditIssueList.SetItemChecked(
                    i,
                    issue != null && issue.AutoApplicable && !issue.Applied
                );
            }
            UpdateAuditActionButtons();
        }

        private bool TryResolveAuditIssueRange(AuditReviewIssue issue, out Word.Range range)
        {
            range = null;
            if (issue == null || string.IsNullOrWhiteSpace(issue.FindText))
                return false;

            Word.Document document;
            Word.Range targetRange;
            if (!TryGetCurrentAuditTarget(out document, out targetRange))
                return false;

            string currentText = targetRange.Text ?? string.Empty;
            int first = currentText.IndexOf(issue.FindText, StringComparison.Ordinal);
            if (first < 0)
                return false;

            int second = currentText.IndexOf(
                issue.FindText,
                first + issue.FindText.Length,
                StringComparison.Ordinal
            );
            if (second >= 0)
                return false;

            range = document.Range(
                targetRange.Start + first,
                targetRange.Start + first + issue.FindText.Length
            );
            return true;
        }

        private void UpdateAuditActionButtons()
        {
            if (_auditApplySelectedButton == null)
                return;

            bool hasSafeChecked = false;
            if (_auditIssueList != null)
            {
                foreach (object checkedItem in _auditIssueList.CheckedItems)
                {
                    AuditReviewIssue issue = checkedItem as AuditReviewIssue;
                    if (issue != null && issue.AutoApplicable && !issue.Applied)
                    {
                        hasSafeChecked = true;
                        break;
                    }
                }
            }

            _auditApplySelectedButton.Enabled = !_auditReviewBusy && hasSafeChecked;
            _auditSelectSafeButton.Enabled = !_auditReviewBusy &&
                _auditReviewIssues.Any(i => i.AutoApplicable && !i.Applied);
        }

        private void SetAuditReviewControlsEnabled(bool enabled)
        {
            if (_auditIssueList != null)
                _auditIssueList.Enabled = enabled;
            if (_auditGoToButton != null)
                _auditGoToButton.Enabled = enabled && _auditIssueList != null && _auditIssueList.SelectedItem != null;
            if (_auditClosePanelButton != null)
                _auditClosePanelButton.Enabled = enabled;
            UpdateAuditActionButtons();
        }

        private void ShowAuditReviewPanel()
        {
            if (_auditReviewPanel == null)
                return;

            if (_evidencePanel != null)
                _evidencePanel.Visible = false;
            _auditReviewPanel.Visible = true;
            _auditReviewPanel.BringToFront();
        }

        private void HideAuditReviewPanel()
        {
            if (_auditReviewPanel != null)
                _auditReviewPanel.Visible = false;
            if (_evidencePanel != null)
                _evidencePanel.Visible = true;
        }

        private void ResetAuditReviewPanelState(bool showEvidence)
        {
            _auditReviewIssues.Clear();
            if (_auditIssueList != null)
                _auditIssueList.Items.Clear();
            ClearAuditPreview();
            if (_auditReviewHeader != null)
                _auditReviewHeader.Text = "Замечания аудита";
            if (_auditReviewPanel != null)
                _auditReviewPanel.Visible = false;
            if (showEvidence && _evidencePanel != null)
                _evidencePanel.Visible = true;
        }
    }
}
