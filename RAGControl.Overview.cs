using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using UglyToad.PdfPig;

namespace TextForge
{
    public partial class RAGControl
    {
        // Generic requests such as "make an overview of the attached PDFs" do not
        // contain a topic that is useful for semantic retrieval. For those requests,
        // read representative pages directly from every checked PDF. This keeps strict
        // RAG grounded in the user's actual files and does not depend on vector-cache
        // health or on an artificial overview query matching the embedding space.
        public List<RagEvidenceItem> GetCheckedPdfOverviewEvidence(int maxTotal = 10)
        {
            var evidence = new List<RagEvidenceItem>();
            maxTotal = Math.Max(1, Math.Min(maxTotal, 10));

            string[] checkedPaths = GetCheckedRagPaths()
                .Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path))
                .OrderBy(path => Path.GetFileName(path), StringComparer.OrdinalIgnoreCase)
                .ToArray();
            if (checkedPaths.Length == 0)
                return evidence;

            var evidenceByFile = new Dictionary<string, List<RagEvidenceItem>>(
                StringComparer.OrdinalIgnoreCase
            );

            foreach (string filePath in checkedPaths)
            {
                var fileEvidence = new List<RagEvidenceItem>();
                try
                {
                    using (PdfDocument document = PdfDocument.Open(filePath))
                    {
                        List<int> pagePlan = BuildRepresentativePagePlan(document.NumberOfPages);
                        foreach (int pageNumber in pagePlan)
                        {
                            string text = ReadOverviewPageText(document, pageNumber, 2600);
                            if (text.Length < 80)
                                continue;

                            fileEvidence.Add(
                                new RagEvidenceItem(
                                    Path.GetFileName(filePath),
                                    pageNumber,
                                    text,
                                    1d
                                )
                            );
                        }
                    }
                }
                catch
                {
                    // A single unreadable PDF must not suppress evidence from the other
                    // checked sources. If every checked PDF is unreadable, strict RAG
                    // will keep its normal empty-evidence guard.
                }

                if (fileEvidence.Count > 0)
                    evidenceByFile[filePath] = fileEvidence;
            }

            // Round-robin page samples so every checked PDF contributes before one file
            // can consume the complete strict-RAG evidence budget.
            int sampleIndex = 0;
            while (evidence.Count < maxTotal)
            {
                bool foundCandidate = false;
                foreach (string filePath in checkedPaths)
                {
                    List<RagEvidenceItem> fileEvidence;
                    if (!evidenceByFile.TryGetValue(filePath, out fileEvidence) ||
                        sampleIndex >= fileEvidence.Count)
                    {
                        continue;
                    }

                    foundCandidate = true;
                    evidence.Add(fileEvidence[sampleIndex]);
                    if (evidence.Count >= maxTotal)
                        break;
                }

                if (!foundCandidate)
                    break;
                sampleIndex++;
            }

            return evidence;
        }

        private static List<int> BuildRepresentativePagePlan(int pageCount)
        {
            var pages = new List<int>();
            if (pageCount <= 0)
                return pages;

            AddRepresentativePage(pages, 1, pageCount);
            AddRepresentativePage(pages, 2, pageCount);
            AddRepresentativePage(pages, (int)Math.Ceiling(pageCount * 0.45d), pageCount);
            AddRepresentativePage(pages, (int)Math.Ceiling(pageCount * 0.70d), pageCount);
            AddRepresentativePage(pages, (int)Math.Ceiling(pageCount * 0.88d), pageCount);
            AddRepresentativePage(pages, pageCount, pageCount);
            return pages;
        }

        private static void AddRepresentativePage(List<int> pages, int pageNumber, int pageCount)
        {
            pageNumber = Math.Max(1, Math.Min(pageNumber, pageCount));
            if (!pages.Contains(pageNumber))
                pages.Add(pageNumber);
        }

        private static string ReadOverviewPageText(
            PdfDocument document,
            int pageNumber,
            int maxCharacters
        )
        {
            var text = new StringBuilder();
            var page = document.GetPage(pageNumber);

            foreach (var word in page.GetWords())
            {
                if (word == null || string.IsNullOrWhiteSpace(word.Text))
                    continue;

                if (text.Length > 0)
                    text.Append(' ');
                text.Append(word.Text.Trim());

                if (text.Length >= maxCharacters)
                    break;
            }

            string result = text.ToString().Trim();
            if (result.Length > maxCharacters)
                result = result.Substring(0, maxCharacters).TrimEnd();
            return result;
        }
    }
}
