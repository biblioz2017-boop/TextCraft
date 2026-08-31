using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    partial class Forge
    {
        private static readonly string[] HelpResourceParts =
        {
            "TextForge.NeZnaikaManualRU.Part01",
            "TextForge.NeZnaikaManualRU.Part02",
            "TextForge.NeZnaikaManualRU.Part03",
            "TextForge.NeZnaikaManualRU.Part04",
            "TextForge.NeZnaikaManualRU.Part05",
            "TextForge.NeZnaikaManualRU.Part06"
        };

        private const string HelpFileName = "NeZnaika-1.0.41-Manual-RU.docx";

        private void HelpButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                OpenEmbeddedManual();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private static void OpenEmbeddedManual()
        {
            Word.Application application = Globals.ThisAddIn.Application;
            string helpDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "NeZnaika",
                "Help"
            );
            Directory.CreateDirectory(helpDirectory);

            string helpPath = Path.Combine(helpDirectory, HelpFileName);
            string fullHelpPath = Path.GetFullPath(helpPath);

            foreach (Word.Document document in application.Documents)
            {
                try
                {
                    if (string.Equals(
                        Path.GetFullPath(document.FullName),
                        fullHelpPath,
                        StringComparison.OrdinalIgnoreCase
                    ))
                    {
                        document.Activate();
                        return;
                    }
                }
                catch
                {
                }
            }

            Assembly assembly = typeof(Forge).Assembly;
            StringBuilder encoded = new StringBuilder(90000);
            foreach (string resourceName in HelpResourceParts)
            {
                using (Stream part = assembly.GetManifestResourceStream(resourceName))
                {
                    if (part == null)
                        throw new InvalidOperationException("Встроенное руководство НеZнайка повреждено: отсутствует ресурс " + resourceName + ".");

                    using (StreamReader reader = new StreamReader(part, Encoding.ASCII, false))
                        encoded.Append(reader.ReadToEnd());
                }
            }

            byte[] manualBytes = Convert.FromBase64String(encoded.ToString());
            File.WriteAllBytes(helpPath, manualBytes);

            application.Documents.Open(
                FileName: helpPath,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: true
            );
        }
    }
}
