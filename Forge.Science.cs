using System;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    partial class Forge
    {
        // Quick rewrite actions use Word's native revision tracking. This gives the user
        // a standard Accept/Reject workflow instead of silently replacing dissertation text.
        private static async Task AnalyzeTextWithTrackChanges(
            string systemPrompt,
            string userPrompt,
            float temperature
        )
        {
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            bool originalTrackRevisions = document.TrackRevisions;

            try
            {
                if (!originalTrackRevisions)
                    document.TrackRevisions = true;

                await AnalyzeText(systemPrompt, userPrompt, temperature);
            }
            finally
            {
                try
                {
                    document.TrackRevisions = originalTrackRevisions;
                }
                catch
                {
                    // Restoring an informational Word setting must not hide the generated text.
                }
            }
        }
    }
}
