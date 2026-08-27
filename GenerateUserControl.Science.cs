using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
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

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            if (_sciencePanelInitialized)
                return;

            _sciencePanelInitialized = true;
            AddScientificAuditTemplate();
            AddScientificQuickActions();
            AddEvidencePanel();

            // The original chat appends tiny streaming fragments and calls ScrollToCaret
            // for every fragment. Batch visual redraws so the text remains stable while
            // still updating several times per second.
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

        private async void GenerateButton_SmoothStreaming(object sender, EventArgs e)
        {
            // The primary async click handler is registered first and disables the button
            // before yielding to streaming. If it did not start, there is nothing to batch.
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

                // One paint for a batch of tokens, then freeze again until the next batch.
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
                    () => rag.GetRAGEvidence(query, 2)
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

            // Common case: one open document but COM returned a different RCW proxy.
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
