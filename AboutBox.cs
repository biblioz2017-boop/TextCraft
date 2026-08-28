using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace TextForge
{
    partial class AboutBox : Form
    {
        private static readonly CultureLocalizationHelper _cultureHelper = new CultureLocalizationHelper("TextForge.AboutBox", typeof(AboutBox).Assembly);
        private RichTextBox _aboutTextBox;

        // User-provided owl-with-globe image, resized for the About dialog and embedded
        // directly in the assembly so the VSTO package remains self-contained/offline.
        private const string NeZnaikaOwlJpegBase64 =
            "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";

        public AboutBox()
        {
            InitializeComponent();

            this.Text = "О программе — НеZнайка";
            this.labelProductName.Text = "НеZнайка";
            this.labelVersion.Text = "Версия " + AssemblyVersion;
            this.labelCopyright.Text = AssemblyCopyright;
            this.labelCompanyName.Text = "Локальная AI-надстройка для Microsoft Word";

            // The original TextCraft About window is too small for the requested text.
            this.Size = new Size(840, 560);
            this.tableLayoutPanel.Dock = DockStyle.Fill;

            TryApplyOwlLogo();
            ConfigureAboutText();
        }

        private void TryApplyOwlLogo()
        {
            try
            {
                byte[] imageBytes = Convert.FromBase64String(NeZnaikaOwlJpegBase64);
                using (MemoryStream stream = new MemoryStream(imageBytes))
                using (Image source = Image.FromStream(stream, true, true))
                {
                    this.logoPictureBox.Image = new Bitmap(source);
                }

                this.logoPictureBox.SizeMode = PictureBoxSizeMode.Zoom;
                this.logoPictureBox.BackColor = Color.White;
            }
            catch
            {
                // A logo problem must never prevent opening the About window.
                this.logoPictureBox.Image = null;
            }
        }

        private void ConfigureAboutText()
        {
            int column = this.tableLayoutPanel.GetColumn(this.LicenseTextBox);
            int row = this.tableLayoutPanel.GetRow(this.LicenseTextBox);
            Font baseFont = this.LicenseTextBox.Font;
            Padding margin = this.LicenseTextBox.Margin;

            this.tableLayoutPanel.Controls.Remove(this.LicenseTextBox);
            this.LicenseTextBox.Dispose();

            _aboutTextBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                DetectUrls = false,
                BackColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle,
                Font = baseFont,
                Margin = margin,
                TabStop = false,
                ScrollBars = RichTextBoxScrollBars.Vertical
            };

            this.tableLayoutPanel.Controls.Add(_aboutTextBox, column, row);

            AppendAboutParagraph("Надстройка, созданная во имя текста, правок и человеческих страданий.", true);
            AppendAboutParagraph("НеZнайка умеет ковырять текст, шаманить над формулировками, совершать обряды форматирования и делать прочие вещи, которые нормальный человек предпочёл бы поручить кому-нибудь другому.", false);
            AppendAboutParagraph("Если документ стал лучше — так и было задумано.", false);
            AppendAboutParagraph("Если стал хуже — это авторский стиль.", false);
            AppendAboutParagraph("Если кончилась оперативка — значит, таинство началось.", false);
            AppendAboutParagraph("Вместе с НеZнайкой мы натянем любую сову на глобус!", true);
            AppendAboutParagraph("И не забывайте страдать!", true);

            _aboutTextBox.AppendText(
                "\r\n────────────────────────────────\r\n" +
                "Сторонние компоненты и лицензии\r\n" +
                "────────────────────────────────\r\n\r\n" +
                Properties.Resources.THIRD_PARTY
            );
            _aboutTextBox.SelectionStart = 0;
            _aboutTextBox.SelectionLength = 0;
        }

        private void AppendAboutParagraph(string text, bool bold)
        {
            _aboutTextBox.SelectionStart = _aboutTextBox.TextLength;
            _aboutTextBox.SelectionLength = 0;
            _aboutTextBox.SelectionFont = new Font(
                _aboutTextBox.Font,
                bold ? FontStyle.Bold : FontStyle.Regular
            );
            _aboutTextBox.AppendText(text + "\r\n\r\n");
        }

        #region Assembly Attribute Accessors

        public string AssemblyTitle
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    if (titleAttribute.Title != "")
                        return titleAttribute.Title;
                }
                return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }

        public string AssemblyVersion
        {
            get { return Assembly.GetExecutingAssembly().GetName().Version.ToString(); }
        }

        public string AssemblyDescription
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                if (attributes.Length == 0) return "";
                return ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }

        public string AssemblyProduct
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                if (attributes.Length == 0) return "";
                return ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        public string AssemblyCopyright
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                if (attributes.Length == 0) return "";
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }

        public string AssemblyCompany
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                if (attributes.Length == 0) return "";
                return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }
        #endregion
    }
}
