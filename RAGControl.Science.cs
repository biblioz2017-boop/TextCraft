using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace TextForge
{
    public partial class RAGControl
    {
        private Label _sourceFilterHintLabel;
        private bool _scienceUiInitialized;

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            if (_scienceUiInitialized)
                return;

            _scienceUiInitialized = true;

            FileListBox.CheckOnClick = true;

            _sourceFilterHintLabel = new Label
            {
                Text = "✓ Отметьте PDF, которые должны участвовать в RAG",
                Dock = DockStyle.Bottom,
                Height = 30,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Padding = new Padding(4, 0, 4, 0)
            };

            Controls.Add(_sourceFilterHintLabel);
            _sourceFilterHintLabel.BringToFront();

            // Existing items are included by default. Newly added files are checked
            // by CheckFileForRag(), called immediately after they are added to the list.
            BeginInvoke((MethodInvoker)delegate
            {
                for (int i = 0; i < FileListBox.Items.Count; i++)
                    FileListBox.SetItemChecked(i, true);
            });
        }

        // Called from the base AddButton handler after BindingList receives a PDF.
        // BeginInvoke waits until the data-bound CheckedListBox has materialized the item.
        private void CheckFileForRag(string filePath)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                for (int i = 0; i < FileListBox.Items.Count; i++)
                {
                    if (FileListBox.Items[i] is KeyValuePair<string, string> file &&
                        string.Equals(file.Value, filePath, StringComparison.OrdinalIgnoreCase))
                    {
                        FileListBox.SetItemChecked(i, true);
                        return;
                    }
                }
            });
        }

        // A checked item participates in retrieval. If no PDF is checked we deliberately
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
            HashSet<string> selected = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            Action capture = () =>
            {
                foreach (object item in FileListBox.CheckedItems)
                {
                    if (item is KeyValuePair<string, string> file)
                        selected.Add(file.Value);
                }
            };

            if (FileListBox.InvokeRequired)
                FileListBox.Invoke((MethodInvoker)delegate { capture(); });
            else
                capture();

            return selected;
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

            return evidence
                .OrderByDescending(item => item.Score)
                .ToList();
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
}
