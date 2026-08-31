using System;
using System.ClientModel;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using OpenAI.Chat;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    partial class Forge
    {
        private sealed class SynonymCandidate
        {
            public string Text { get; set; }
            public string Register { get; set; }
            public string Note { get; set; }
        }

        private async void SynonymsButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Word.Selection selection = Globals.ThisAddIn.Application.Selection;
                if (selection == null || selection.End <= selection.Start)
                {
                    MessageBox.Show(
                        "Сначала выделите слово или короткую фразу, для которой нужно подобрать синонимы.",
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

                if (selectedText.Length > 220)
                {
                    MessageBox.Show(
                        "Для подбора синонимов выделите слово или короткую фразу длиной до 220 символов.",
                        "НеZнайка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    return;
                }

                Word.Range contextRange = sourceRange.Duplicate;
                try
                {
                    contextRange.Expand(Word.WdUnits.wdSentence);
                }
                catch
                {
                    contextRange = sourceRange.Paragraphs.Count > 0
                        ? sourceRange.Paragraphs[1].Range.Duplicate
                        : sourceRange.Duplicate;
                }

                string context = CommonUtils.SubstringTokens(
                    contextRange.Text ?? selectedText,
                    Math.Max(256, (int)(ThisAddIn.ContextLength * 0.08))
                );

                SetStatus("◌ Подбираю синонимы…");
                CancelButtonVisibility(true);
                List<SynonymCandidate> candidates = await GetContextualSynonymsAsync(
                    selectedText,
                    context
                );

                if (candidates.Count == 0)
                {
                    MessageBox.Show(
                        "Модель не нашла безопасной контекстной замены. Исходный текст оставлен без изменений.",
                        "НеZнайка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    return;
                }

                using (var dialog = new SynonymPickerForm(selectedText, context, candidates))
                {
                    if (dialog.ShowDialog() != DialogResult.OK || string.IsNullOrWhiteSpace(dialog.SelectedSynonym))
                        return;

                    ReplaceSelectionWithSynonym(sourceRange, dialog.SelectedSynonym);
                }
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

        private static async Task<List<SynonymCandidate>> GetContextualSynonymsAsync(
            string selectedText,
            string context
        )
        {
            ChatClient client = new ChatClient(
                ThisAddIn.Model,
                new ApiKeyCredential(ThisAddIn.ApiKey),
                ThisAddIn.ClientOptions
            );

            var messages = new List<ChatMessage>()
            {
                new SystemChatMessage(
                    "Ты редактор русского научного текста. Подбирай синонимы и близкие по смыслу замены только с учётом контекста. " +
                    "Не предлагай вариант, если он меняет фактический смысл, степень уверенности, причинность, терминологическое значение или научную точность. " +
                    "Для специального термина допускается ответ НЕТ БЕЗОПАСНОЙ ЗАМЕНЫ."
                ),
                new UserChatMessage(
                    "Выделение: " + selectedText + "\n\n" +
                    "Контекст предложения:\n" + context + "\n\n" +
                    "Предложи 5–8 контекстно допустимых замен. Для каждой укажи оттенок и очень коротко объясни различие. " +
                    "Формат каждой строки строго: ВАРИАНТ<TAB>МЕТКА<TAB>ПОЯСНЕНИЕ. " +
                    "Метки используй из набора: нейтральный; научный; более точный; формальный; мягче; сильнее. " +
                    "Не используй Markdown, нумерацию и дополнительные строки. Если безопасных замен нет, верни ровно: НЕТ БЕЗОПАСНОЙ ЗАМЕНЫ"
                )
            };

            var stream = client.CompleteChatStreamingAsync(
                messages,
                new ChatCompletionOptions() { Temperature = 0.28f },
                ThisAddIn.CancellationTokenSource.Token
            );

            StringBuilder buffer = new StringBuilder();
            await foreach (
                var update in stream.WithCancellation(ThisAddIn.CancellationTokenSource.Token)
            )
            {
                foreach (var part in update.ContentUpdate)
                {
                    if (part.Kind == ChatMessageContentPartKind.Text)
                        buffer.Append(part.Text);
                }
            }

            return ParseSynonymCandidates(buffer.ToString(), selectedText);
        }

        private static List<SynonymCandidate> ParseSynonymCandidates(string raw, string selectedText)
        {
            var result = new List<SynonymCandidate>();
            var seen = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
            string text = (raw ?? string.Empty).Trim();
            if (text.IndexOf("НЕТ БЕЗОПАСНОЙ ЗАМЕНЫ", StringComparison.OrdinalIgnoreCase) >= 0)
                return result;

            string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            foreach (string sourceLine in lines)
            {
                string line = sourceLine.Trim().TrimStart('-', '•', '*').Trim();
                if (line.Length == 0)
                    continue;

                string[] parts = line.Split(new[] { '\t' }, 3);
                if (parts.Length < 2)
                    parts = line.Split(new[] { " | " }, 3, StringSplitOptions.None);
                if (parts.Length < 2)
                    continue;

                string candidate = parts[0].Trim().Trim('"', '«', '»');
                if (candidate.Length == 0 || candidate.Length > 220)
                    continue;
                if (string.Equals(candidate, selectedText, StringComparison.CurrentCultureIgnoreCase))
                    continue;
                if (!seen.Add(candidate))
                    continue;

                result.Add(new SynonymCandidate
                {
                    Text = candidate,
                    Register = parts[1].Trim(),
                    Note = parts.Length > 2 ? parts[2].Trim() : string.Empty
                });

                if (result.Count >= 8)
                    break;
            }

            return result;
        }

        private static void ReplaceSelectionWithSynonym(Word.Range originalRange, string replacement)
        {
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            bool originalTrackRevisions = document.TrackRevisions;
            Word.Range range = originalRange.Duplicate;

            try
            {
                if (!originalTrackRevisions)
                    document.TrackRevisions = true;

                range.Text = replacement;
                Globals.ThisAddIn.Application.Selection.SetRange(range.Start, range.Start + replacement.Length);
            }
            finally
            {
                try { document.TrackRevisions = originalTrackRevisions; } catch { }
            }
        }

        private sealed class SynonymPickerForm : Form
        {
            private readonly DataGridView _grid;
            public string SelectedSynonym { get; private set; }

            public SynonymPickerForm(
                string source,
                string context,
                List<SynonymCandidate> candidates
            )
            {
                Text = "НеZнайка — синонимы";
                StartPosition = FormStartPosition.CenterParent;
                Width = 780;
                Height = 430;
                MinimizeBox = false;
                MaximizeBox = false;
                ShowInTaskbar = false;

                var sourceLabel = new Label
                {
                    Dock = DockStyle.Top,
                    Height = 34,
                    Padding = new Padding(10, 9, 10, 0),
                    Text = "Выделение: " + source
                };

                var contextBox = new TextBox
                {
                    Dock = DockStyle.Top,
                    Height = 64,
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Vertical,
                    Text = context,
                    BackColor = System.Drawing.SystemColors.Window
                };

                _grid = new DataGridView
                {
                    Dock = DockStyle.Fill,
                    ReadOnly = true,
                    AllowUserToAddRows = false,
                    AllowUserToDeleteRows = false,
                    AllowUserToResizeRows = false,
                    AutoGenerateColumns = false,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                    MultiSelect = false,
                    RowHeadersVisible = false
                };
                _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Вариант", DataPropertyName = "Text", Width = 190 });
                _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Оттенок", DataPropertyName = "Register", Width = 120 });
                _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Почему подходит", DataPropertyName = "Note", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
                _grid.DataSource = candidates;
                _grid.CellDoubleClick += (s, e) => AcceptCurrent();

                var buttons = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    Height = 46,
                    FlowDirection = FlowDirection.RightToLeft,
                    Padding = new Padding(8)
                };
                var cancel = new Button { Text = "Отмена", AutoSize = true, DialogResult = DialogResult.Cancel };
                var replace = new Button { Text = "Заменить", AutoSize = true };
                replace.Click += (s, e) => AcceptCurrent();
                buttons.Controls.Add(cancel);
                buttons.Controls.Add(replace);

                Controls.Add(_grid);
                Controls.Add(contextBox);
                Controls.Add(sourceLabel);
                Controls.Add(buttons);

                AcceptButton = replace;
                CancelButton = cancel;
            }

            private void AcceptCurrent()
            {
                if (_grid.CurrentRow == null)
                    return;
                var item = _grid.CurrentRow.DataBoundItem as SynonymCandidate;
                if (item == null || string.IsNullOrWhiteSpace(item.Text))
                    return;

                SelectedSynonym = item.Text;
                DialogResult = DialogResult.OK;
                Close();
            }
        }
    }
}
