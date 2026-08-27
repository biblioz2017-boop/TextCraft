using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    public partial class RAGControl
    {
        private Label _sourceFilterHintLabel;
        private Label _sourceCountLabel;
        private FlowLayoutPanel _sourceActionsPanel;
        private Button _checkAllSourcesButton;
        private Button _uncheckAllSourcesButton;
        private Button _invertSourcesButton;
        private bool _scienceUiInitialized;
        private Timer _sourceBindingTimer;
        private bool _sourceListHooked;
        private bool _applyingChecks;
        private readonly object _includedSourcesLock = new object();
        private readonly HashSet<string> _includedSourcePaths =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            if (_scienceUiInitialized)
                return;

            _scienceUiInitialized = true;
            FileListBox.CheckOnClick = true;
            FileListBox.ItemCheck += FileListBox_ItemCheck;

            _sourceFilterHintLabel = new Label
            {
                Text = "✓ Галочка = источник участвует в RAG",
                Dock = DockStyle.Bottom,
                Height = 24,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Padding = new Padding(4, 0, 4, 0)
            };

            _sourceActionsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 58,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Padding = new Padding(3, 2, 3, 2)
            };

            _checkAllSourcesButton = new Button
            {
                Text = "Все",
                AutoSize = true,
                Height = 25,
                Margin = new Padding(1)
            };
            _checkAllSourcesButton.Click += (s, args) => SetAllSourceChecks(true);

            _uncheckAllSourcesButton = new Button
            {
                Text = "Снять все",
                AutoSize = true,
                Height = 25,
                Margin = new Padding(1)
            };
            _uncheckAllSourcesButton.Click += (s, args) => SetAllSourceChecks(false);

            _invertSourcesButton = new Button
            {
                Text = "Инвертировать",
                AutoSize = true,
                Height = 25,
                Margin = new Padding(1)
            };
            _invertSourcesButton.Click += (s, args) => InvertSourceChecks();

            _sourceCountLabel = new Label
            {
                Text = "В RAG: 0 из 0",
                AutoSize = true,
                Height = 25,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Padding = new Padding(5, 5, 0, 0),
                Margin = new Padding(1)
            };

            _sourceActionsPanel.Controls.Add(_checkAllSourcesButton);
            _sourceActionsPanel.Controls.Add(_uncheckAllSourcesButton);
            _sourceActionsPanel.Controls.Add(_invertSourcesButton);
            _sourceActionsPanel.Controls.Add(_sourceCountLabel);

            Controls.Add(_sourceActionsPanel);
            Controls.Add(_sourceFilterHintLabel);
            _sourceActionsPanel.BringToFront();
            _sourceFilterHintLabel.BringToFront();

            // _fileList is initialized asynchronously by the original control. Poll for
            // it briefly, then subscribe once. Inclusion state is stored by the real PDF
            // path, not by the visual row, so label changes such as [CACHE] -> [OK] do not
            // silently re-enable a source the user explicitly unchecked.
            _sourceBindingTimer = new Timer { Interval = 150 };
            _sourceBindingTimer.Tick += (s, args) =>
            {
                if (_sourceListHooked || _fileList == null)
                    return;

                _sourceListHooked = true;
                _sourceBindingTimer.Stop();
                _sourceBindingTimer.Dispose();

                lock (_includedSourcesLock)
                {
                    foreach (KeyValuePair<string, string> file in _fileList)
                        _includedSourcePaths.Add(file.Value);
                }

                _fileList.ListChanged += (ls, le) =>
                {
                    if (le.ListChangedType == System.ComponentModel.ListChangedType.ItemAdded &&
                        le.NewIndex >= 0 && le.NewIndex < _fileList.Count)
                    {
                        string addedPath = _fileList[le.NewIndex].Value;
                        lock (_includedSourcesLock)
                            _includedSourcePaths.Add(addedPath);
                    }

                    // BindingList row replacement is used for [INDEXING]/[CACHE]/[OK]
                    // statuses. Reapply the path-based inclusion state after any change.
                    BeginInvoke((MethodInvoker)delegate
                    {
                        ApplyStoredChecksToList();
                        UpdateSourceCountLabel();
                    });
                };

                BeginInvoke((MethodInvoker)delegate
                {
                    ApplyStoredChecksToList();
                    UpdateSourceCountLabel();
                });
            };
            _sourceBindingTimer.Start();
        }

        private void FileListBox_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (_applyingChecks || e.Index < 0 || e.Index >= FileListBox.Items.Count)
                return;

            if (FileListBox.Items[e.Index] is KeyValuePair<string, string> file)
            {
                lock (_includedSourcesLock)
                {
                    if (e.NewValue == CheckState.Checked)
                        _includedSourcePaths.Add(file.Value);
                    else
                        _includedSourcePaths.Remove(file.Value);
                }
            }

            // ItemCheck fires before CheckedItems has changed; defer the visual count.
            BeginInvoke((MethodInvoker)delegate { UpdateSourceCountLabel(); });
        }

        private void SetAllSourceChecks(bool isChecked)
        {
            _applyingChecks = true;
            try
            {
                lock (_includedSourcesLock)
                {
                    if (!isChecked)
                    {
                        _includedSourcePaths.Clear();
                    }
                    else
                    {
                        foreach (object item in FileListBox.Items)
                        {
                            if (item is KeyValuePair<string, string> file)
                                _includedSourcePaths.Add(file.Value);
                        }
                    }
                }

                for (int i = 0; i < FileListBox.Items.Count; i++)
                    FileListBox.SetItemChecked(i, isChecked);
            }
            finally
            {
                _applyingChecks = false;
            }

            UpdateSourceCountLabel();
        }

        private void InvertSourceChecks()
        {
            _applyingChecks = true;
            try
            {
                for (int i = 0; i < FileListBox.Items.Count; i++)
                {
                    bool newState = !FileListBox.GetItemChecked(i);
                    if (FileListBox.Items[i] is KeyValuePair<string, string> file)
                    {
                        lock (_includedSourcesLock)
                        {
                            if (newState)
                                _includedSourcePaths.Add(file.Value);
                            else
                                _includedSourcePaths.Remove(file.Value);
                        }
                    }
                    FileListBox.SetItemChecked(i, newState);
                }
            }
            finally
            {
                _applyingChecks = false;
            }

            UpdateSourceCountLabel();
        }

        private void ApplyStoredChecksToList()
        {
            if (FileListBox == null || FileListBox.IsDisposed)
                return;

            _applyingChecks = true;
            try
            {
                for (int i = 0; i < FileListBox.Items.Count; i++)
                {
                    if (!(FileListBox.Items[i] is KeyValuePair<string, string> file))
                        continue;

                    bool isIncluded;
                    lock (_includedSourcesLock)
                        isIncluded = _includedSourcePaths.Contains(file.Value);

                    if (FileListBox.GetItemChecked(i) != isIncluded)
                        FileListBox.SetItemChecked(i, isIncluded);
                }
            }
            finally
            {
                _applyingChecks = false;
            }
        }

        private void UpdateSourceCountLabel()
        {
            if (_sourceCountLabel == null || _sourceCountLabel.IsDisposed)
                return;

            int total = FileListBox == null ? 0 : FileListBox.Items.Count;
            int included = 0;

            lock (_includedSourcesLock)
            {
                if (_fileList != null)
                    included = _fileList.Count(file => _includedSourcePaths.Contains(file.Value));
            }

            _sourceCountLabel.Text = "В RAG: " + included + " из " + total;
            _checkAllSourcesButton.Enabled = total > 0 && included < total;
            _uncheckAllSourcesButton.Enabled = included > 0;
            _invertSourcesButton.Enabled = total > 0;
        }

        // A checked source participates in retrieval. If no PDF is checked we deliberately
        // return an empty set rather than silently falling back to all sources.
        private KeyValuePair<string, HyperVectorDB.HyperVectorDB>[] GetActiveRagDatabases()
        {
            KeyValuePair<string, HyperVectorDB.HyperVectorDB>[] all = _fileDatabases.ToArray();
            if (all.Length == 0)
                return all;

            HashSet<string> checkedPaths = GetCheckedRagPaths();
            if (checkedPaths.Count == 0)
                return Array.Empty<KeyValuePair<string, HyperVectorDB.HyperVectorDB>>();

            return all
                .Where(item => checkedPaths.Contains(item.Key))
                .ToArray();
        }

        private HashSet<string> GetCheckedRagPaths()
        {
            lock (_includedSourcesLock)
                return new HashSet<string>(_includedSourcePaths, StringComparer.OrdinalIgnoreCase);
        }

        public List<RagEvidenceItem> GetRAGEvidence(string query, int maxPerFile = 2)
        {
            List<RagEvidenceItem> evidence = new List<RagEvidenceItem>();
            if (string.IsNullOrWhiteSpace(query))
                return evidence;

            maxPerFile = Math.Max(1, Math.Min(maxPerFile, 6));

            foreach (var entry in GetActiveRagDatabases().OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase))
            {
                var result = entry.Value.QueryCosineSimilarity(query, maxPerFile);
                for (int i = 0; i < result.Documents.Count; i++)
                {
                    string raw = result.Documents[i].DocumentString ?? string.Empty;
                    double score = i < result.Distances.Count ? result.Distances[i] : 0d;
                    evidence.Add(ParseEvidence(entry.Key, raw, score));
                }
            }

            return evidence.OrderByDescending(item => item.Score).ToList();
        }

        private static RagEvidenceItem ParseEvidence(string filePath, string raw, double score)
        {
            string source = Path.GetFileName(filePath);
            int? page = null;
            string text = raw.Trim();

            Match metadata = Regex.Match(
                raw,
                @"^\[Source:\s*(.*?);\s*Page:\s*(\d+)\]\s*(.*)$",
                RegexOptions.Singleline | RegexOptions.IgnoreCase
            );

            if (metadata.Success)
            {
                source = metadata.Groups[1].Value.Trim();
                if (int.TryParse(metadata.Groups[2].Value, out int parsedPage))
                    page = parsedPage;
                text = metadata.Groups[3].Value.Trim();
            }

            return new RagEvidenceItem(source, page, text, score);
        }

        private static string GetPersistentDatabasePath(string filePath)
        {
            FileInfo info = new FileInfo(filePath);
            string identity =
                Path.GetFullPath(filePath) + "|" +
                info.Length + "|" +
                info.LastWriteTimeUtc.Ticks + "|" +
                (ThisAddIn.EmbedModel ?? string.Empty) + "|" +
                CHUNK_LEN + "|" + CHUNK_OVERLAP;

            byte[] hash;
            using (SHA256 sha = SHA256.Create())
                hash = sha.ComputeHash(Encoding.UTF8.GetBytes(identity));

            StringBuilder hex = new StringBuilder(hash.Length * 2);
            foreach (byte b in hash)
                hex.Append(b.ToString("x2"));

            string root = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "TextCraft",
                "RAGCache"
            );

            Directory.CreateDirectory(root);
            return Path.Combine(root, hex.ToString());
        }

        private bool TryLoadCachedDatabase(string filePath, out HyperVectorDB.HyperVectorDB database)
        {
            database = null;
            try
            {
                string cachePath = GetPersistentDatabasePath(filePath);
                string indexFile = Path.Combine(cachePath, "indexs.txt");
                if (!File.Exists(indexFile))
                    return false;

                var cached = new HyperVectorDB.HyperVectorDB(ThisAddIn.Embedder, cachePath);
                cached.Load();
                database = cached;
                return true;
            }
            catch
            {
                database = null;
                return false;
            }
        }

        public sealed class RagEvidenceItem
        {
            public string Source { get; private set; }
            public int? Page { get; private set; }
            public string Text { get; private set; }
            public double Score { get; private set; }

            public RagEvidenceItem(string source, int? page, string text, double score)
            {
                Source = source;
                Page = page;
                Text = text;
                Score = score;
            }

            public string CitationLabel
            {
                get
                {
                    return Page.HasValue
                        ? "[" + Source + ", с. " + Page.Value + "]"
                        : "[" + Source + "]";
                }
            }
        }
    }

    // Word can hand VSTO different RCW objects for the same underlying COM document.
    // The default Dictionary comparer treats those wrappers as different keys, which
    // caused ProcessInformation() to throw KeyNotFoundException even though the task
    // panes for that document were already present. Compare by COM IUnknown identity
    // instead, so every RCW for the same Word.Document resolves to the same entry.
    public partial class ThisAddIn
    {
        static ThisAddIn()
        {
            _allTaskPanes = new Dictionary<
                Word.Document,
                Tuple<Microsoft.Office.Tools.CustomTaskPane, Microsoft.Office.Tools.CustomTaskPane, RAGControl>
            >(new WordDocumentComIdentityComparer());
        }

        private sealed class WordDocumentComIdentityComparer : IEqualityComparer<Word.Document>
        {
            public bool Equals(Word.Document x, Word.Document y)
            {
                if (ReferenceEquals(x, y))
                    return true;
                if (x == null || y == null)
                    return false;

                IntPtr xIdentity = IntPtr.Zero;
                IntPtr yIdentity = IntPtr.Zero;
                try
                {
                    xIdentity = Marshal.GetIUnknownForObject(x);
                    yIdentity = Marshal.GetIUnknownForObject(y);
                    return xIdentity == yIdentity;
                }
                catch
                {
                    return false;
                }
                finally
                {
                    if (xIdentity != IntPtr.Zero)
                        Marshal.Release(xIdentity);
                    if (yIdentity != IntPtr.Zero)
                        Marshal.Release(yIdentity);
                }
            }

            public int GetHashCode(Word.Document obj)
            {
                if (obj == null)
                    return 0;

                IntPtr identity = IntPtr.Zero;
                try
                {
                    identity = Marshal.GetIUnknownForObject(obj);
                    return identity.GetHashCode();
                }
                catch
                {
                    return RuntimeHelpers.GetHashCode(obj);
                }
                finally
                {
                    if (identity != IntPtr.Zero)
                        Marshal.Release(identity);
                }
            }
        }
    }
}
