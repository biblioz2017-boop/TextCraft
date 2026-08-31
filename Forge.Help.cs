using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    partial class Forge
    {
        private const string HelpResourceName = "TextForge.NeZnaikaManualRU.docx";
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
            using (Stream source = assembly.GetManifestResourceStream(HelpResourceName))
            {
                if (source == null)
                    throw new InvalidOperationException("Встроенное руководство НеZнайка не найдено в ресурсах надстройки.");

                using (FileStream target = new FileStream(
                    helpPath,
                    FileMode.Create,
                    FileAccess.Write,
                    FileShare.Read
                ))
                {
                    source.CopyTo(target);
                }
            }

            application.Documents.Open(
                FileName: helpPath,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: true
            );
        }
    }
}
