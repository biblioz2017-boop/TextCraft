using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    public partial class GenerateUserControl
    {
        private Panel _evidencePanel;
        private RichTextBox _evidenceTextBox;
        private Button _auditChapterButton;
        private bool _sciencePanelInitialized;
        private int _evidenceRequestVersion;

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            if (_sciencePanelInitialized)
                return;

            _sciencePanelInitialized = true;
            AddScientificAuditTemplate();
            AddScientificQuickActions();
            AddEvidencePanel();

            // The main async click handler is registered first by the designer. This
            // secondary handler captures the query and refreshes evidence after generation.
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
                Text = "Доказательства из выбранных PDF:",
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

            // The primary async handler disables the button while the model streams.
            // Wait until it finishes so evidence and the visible answer belong to one query.
            for (int i = 0; i < 120 && !GenerateButton.Enabled; i++)
                await Task.Delay(250);

            if (requestVersion != _evidenceRequestVersion)
                return;

            try
            {
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                if (!ThisAddIn.AllTaskPanes.TryGetValue(doc, out var panes))
                {
                    _evidenceTextBox.Text = "RAG-панель для документа не найдена.";
                    return;
                }

                RAGControl rag = panes.Item3;
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

        private void RenderEvidence(List<RAGControl.RagEvidenceItem> evidence)
        {
            if (evidence == null || evidence.Count == 0)
            {
                _evidenceTextBox.Text =
                    "Подходящие фрагменты в активных PDF не найдены. " +
                    "Если включен режим 'только выделенные PDF', проверьте выбор источников.";
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
