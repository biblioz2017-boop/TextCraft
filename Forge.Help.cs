using System;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace TextForge
{
    partial class Forge
    {
        private static readonly string[] HelpResourceParts =
        {
            "TextForge.NeZnaikaManualRU.GzipPart01",
            "TextForge.NeZnaikaManualRU.GzipPart02",
            "TextForge.NeZnaikaManualRU.GzipPart03",
            "TextForge.NeZnaikaManualRU.GzipPart04"
        };

        private Form _helpForm;

        private void HelpButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_helpForm != null && !_helpForm.IsDisposed)
                {
                    if (_helpForm.WindowState == FormWindowState.Minimized)
                        _helpForm.WindowState = FormWindowState.Normal;
                    _helpForm.BringToFront();
                    _helpForm.Activate();
                    return;
                }

                string manualText = ReadEmbeddedHelpText();
                _helpForm = CreateEmbeddedHelpForm(manualText);
                _helpForm.FormClosed += (s, args) => _helpForm = null;
                _helpForm.Show();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private static string ReadEmbeddedHelpText()
        {
            Assembly assembly = typeof(Forge).Assembly;
            StringBuilder encoded = new StringBuilder(24000);

            foreach (string resourceName in HelpResourceParts)
            {
                using (Stream part = assembly.GetManifestResourceStream(resourceName))
                {
                    if (part == null)
                    {
                        throw new InvalidOperationException(
                            "Встроенное руководство НеZнайка повреждено: отсутствует ресурс " + resourceName + "."
                        );
                    }

                    using (StreamReader reader = new StreamReader(part, Encoding.ASCII, false))
                        encoded.Append(reader.ReadToEnd());
                }
            }

            byte[] compressed = Convert.FromBase64String(encoded.ToString());
            using (MemoryStream input = new MemoryStream(compressed))
            using (GZipStream gzip = new GZipStream(input, CompressionMode.Decompress))
            using (StreamReader reader = new StreamReader(gzip, Encoding.UTF8, true))
                return reader.ReadToEnd();
        }

        private static Form CreateEmbeddedHelpForm(string manualText)
        {
            Form form = new Form
            {
                Text = "НеZнайка — руководство 1.0.41",
                StartPosition = FormStartPosition.CenterScreen,
                Width = 920,
                Height = 760,
                MinimumSize = new Size(650, 480),
                ShowInTaskbar = true,
                KeyPreview = true
            };

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(8)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            FlowLayoutPanel searchBar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0, 0, 0, 6)
            };

            Label searchLabel = new Label
            {
                Text = "Поиск:",
                AutoSize = true,
                Margin = new Padding(0, 7, 6, 0)
            };

            TextBox searchBox = new TextBox
            {
                Width = 470,
                Margin = new Padding(0, 3, 6, 0)
            };

            Button findNextButton = new Button
            {
                Text = "Найти далее",
                AutoSize = true,
                Margin = new Padding(0, 2, 6, 0)
            };

            Button topButton = new Button
            {
                Text = "В начало",
                AutoSize = true,
                Margin = new Padding(0, 2, 0, 0)
            };

            RichTextBox textBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                DetectUrls = true,
                BackColor = SystemColors.Window,
                Font = new Font("Segoe UI", 10f),
                Text = manualText ?? string.Empty,
                ScrollBars = RichTextBoxScrollBars.ForcedVertical,
                WordWrap = true,
                HideSelection = false
            };

            Label status = new Label
            {
                Text = "Встроенное руководство НеZнайка 1.0.41",
                Dock = DockStyle.Fill,
                AutoSize = true,
                Padding = new Padding(0, 6, 0, 0)
            };

            Action findNext = () =>
            {
                string query = (searchBox.Text ?? string.Empty).Trim();
                if (query.Length == 0)
                {
                    searchBox.Focus();
                    return;
                }

                int start = Math.Min(textBox.TextLength, Math.Max(0, textBox.SelectionStart + textBox.SelectionLength));
                int found = textBox.Find(query, start, RichTextBoxFinds.None);
                if (found < 0 && start > 0)
                    found = textBox.Find(query, 0, RichTextBoxFinds.None);

                if (found < 0)
                {
                    status.Text = "Не найдено: " + query;
                    return;
                }

                textBox.Select(found, query.Length);
                textBox.ScrollToCaret();
                textBox.Focus();
                status.Text = "Найдено: " + query;
            };

            findNextButton.Click += (s, e) => findNext();
            searchBox.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.SuppressKeyPress = true;
                    findNext();
                }
            };
            topButton.Click += (s, e) =>
            {
                textBox.Select(0, 0);
                textBox.ScrollToCaret();
                textBox.Focus();
                status.Text = "Встроенное руководство НеZнайка 1.0.41";
            };
            form.KeyDown += (s, e) =>
            {
                if (e.Control && e.KeyCode == Keys.F)
                {
                    e.SuppressKeyPress = true;
                    searchBox.Focus();
                    searchBox.SelectAll();
                }
                else if (e.KeyCode == Keys.Escape)
                {
                    form.Close();
                }
            };

            searchBar.Controls.Add(searchLabel);
            searchBar.Controls.Add(searchBox);
            searchBar.Controls.Add(findNextButton);
            searchBar.Controls.Add(topButton);

            layout.Controls.Add(searchBar, 0, 0);
            layout.Controls.Add(textBox, 0, 1);
            layout.Controls.Add(status, 0, 2);
            form.Controls.Add(layout);

            form.Shown += (s, e) =>
            {
                textBox.Select(0, 0);
                textBox.ScrollToCaret();
            };

            return form;
        }
    }
}
