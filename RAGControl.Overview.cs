using System;
using System.Collections.Generic;
using System.Linq;

namespace TextForge
{
    public partial class RAGControl
    {
        // Generic requests such as "make an overview of the attached PDFs" do not
        // contain a scientific topic that is meaningful for semantic retrieval.
        // Instead, collect representative indexed chunks from every checked PDF.
        // The caller still applies the normal strict-RAG citation guard, so these
        // chunks remain the only external evidence available to the model.
        public List<RagEvidenceItem> GetCheckedPdfOverviewEvidence(int maxTotal = 10)
        {
            var evidence = new List<RagEvidenceItem>();
            maxTotal = Math.Max(1, Math.Min(maxTotal, 10));

            var databases = GetActiveRagDatabases()
                .OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            if (databases.Length == 0)
                return evidence;

            string[] overviewQueries =
            {
                "summary abstract introduction purpose objective overview " +
                    "\u0430\u043d\u043d\u043e\u0442\u0430\u0446\u0438\u044f \u0432\u0432\u0435\u0434\u0435\u043d\u0438\u0435 \u0446\u0435\u043b\u044c \u043e\u0431\u0437\u043e\u0440",
                "methods results findings discussion analysis " +
                    "\u043c\u0435\u0442\u043e\u0434\u044b \u0440\u0435\u0437\u0443\u043b\u044c\u0442\u0430\u0442\u044b \u043e\u0431\u0441\u0443\u0436\u0434\u0435\u043d\u0438\u0435 \u0430\u043d\u0430\u043b\u0438\u0437",
                "conclusion implications limitations key findings " +
                    "\u0437\u0430\u043a\u043b\u044e\u0447\u0435\u043d\u0438\u0435 \u0432\u044b\u0432\u043e\u0434\u044b \u043e\u0433\u0440\u0430\u043d\u0438\u0447\u0435\u043d\u0438\u044f \u0437\u043d\u0430\u0447\u0435\u043d\u0438\u0435"
            };

            var seenByFile = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            var countByFile = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in databases)
            {
                seenByFile[entry.Key] = new HashSet<string>(StringComparer.Ordinal);
                countByFile[entry.Key] = 0;
            }

            // Round-robin by section type so every checked PDF contributes before any
            // single paper can consume the complete ten-fragment strict-RAG budget.
            foreach (string overviewQuery in overviewQueries)
            {
                foreach (var entry in databases)
                {
                    if (evidence.Count >= maxTotal)
                        break;

                    try
                    {
                        var result = entry.Value.QueryCosineSimilarity(overviewQuery, 2);
                        for (int i = 0; i < result.Documents.Count; i++)
                        {
                            string raw = result.Documents[i].DocumentString ?? string.Empty;
                            if (string.IsNullOrWhiteSpace(raw) || !seenByFile[entry.Key].Add(raw))
                                continue;

                            double score = i < result.Distances.Count ? result.Distances[i] : 0d;
                            evidence.Add(ParseEvidence(entry.Key, raw, score));
                            countByFile[entry.Key]++;
                            break;
                        }
                    }
                    catch
                    {
                        // One damaged/stale index must not suppress evidence from the
                        // other checked PDFs. The strict caller will still reject an
                        // entirely empty evidence set.
                    }
                }

                if (evidence.Count >= maxTotal)
                    break;
            }

            // If a PDF did not match the structural queries above, make one neutral
            // fallback query so that a checked source still gets a chance to appear in
            // the overview. This remains vector-index evidence, not model knowledge.
            const string fallbackQuery =
                "research study paper document main content scientific text " +
                "\u0438\u0441\u0441\u043b\u0435\u0434\u043e\u0432\u0430\u043d\u0438\u0435 \u0441\u0442\u0430\u0442\u044c\u044f \u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442 \u043e\u0441\u043d\u043e\u0432\u043d\u043e\u0435 \u0441\u043e\u0434\u0435\u0440\u0436\u0430\u043d\u0438\u0435";

            foreach (var entry in databases)
            {
                if (evidence.Count >= maxTotal)
                    break;
                if (countByFile[entry.Key] > 0)
                    continue;

                try
                {
                    var result = entry.Value.QueryCosineSimilarity(fallbackQuery, 3);
                    for (int i = 0; i < result.Documents.Count; i++)
                    {
                        string raw = result.Documents[i].DocumentString ?? string.Empty;
                        if (string.IsNullOrWhiteSpace(raw) || !seenByFile[entry.Key].Add(raw))
                            continue;

                        double score = i < result.Distances.Count ? result.Distances[i] : 0d;
                        evidence.Add(ParseEvidence(entry.Key, raw, score));
                        countByFile[entry.Key]++;
                        break;
                    }
                }
                catch
                {
                }
            }

            return evidence;
        }
    }
}
