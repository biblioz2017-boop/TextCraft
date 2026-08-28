using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace TextForge
{
    partial class AboutBox : Form
    {
        private static readonly CultureLocalizationHelper _cultureHelper =
            new CultureLocalizationHelper("TextForge.AboutBox", typeof(AboutBox).Assembly);

        private const string OwlResourceName = "TextForge.NeZnaikaOwl.png";

        private const string NeZnaikaAboutText =
            "Надстройка, созданная во имя текста, правок и человеческих страданий.\r\n\r\n" +
            "НеZнайка умеет ковырять текст, шаманить над формулировками, совершать обряды форматирования и делать прочие вещи, которые нормальный человек предпочёл бы поручить кому-нибудь другому.\r\n\r\n" +
            "Если документ стал лучше — так и было задумано.\r\n\r\n" +
            "Если стал хуже — это авторский стиль.\r\n\r\n" +
            "Если кончилась оперативка — значит, таинство началось.\r\n\r\n" +
            "Вместе с НеZнайкой мы натянем любую сову на глобус!\r\n\r\n" +
            "И не забывайте страдать!";

        public AboutBox()
        {
            InitializeComponent();

            this.Text = "О программе — НеZнайка";
            this.labelProductName.Text = "НеZнайка";
            this.labelVersion.Text = "Версия " + AssemblyVersion;
            this.labelCopyright.Text = AssemblyCopyright;
            this.labelCompanyName.Text = "Локальная AI-надстройка для Microsoft Word";

            // Do not store System.Drawing.Image objects in AboutBox.resx. MSBuild's
            // GenerateResource serializes those objects through GDI+ and can fail on
            // headless Windows runners. The PNG is embedded as a raw manifest resource
            // and decoded only when this dialog is actually opened in Word.
            this.logoPictureBox.SizeMode = PictureBoxSizeMode.Zoom;
            this.logoPictureBox.BackColor = Color.White;
            LoadNeZnaikaLogo();

            this.Size = new Size(900, 620);
            this.tableLayoutPanel.Dock = DockStyle.Fill;

            this.LicenseTextBox.Multiline = true;
            this.LicenseTextBox.ScrollBars = ScrollBars.Vertical;
            this.LicenseTextBox.WordWrap = true;
            this.LicenseTextBox.ReadOnly = true;
            this.LicenseTextBox.BackColor = SystemColors.Window;
            this.LicenseTextBox.Text =
                NeZnaikaAboutText +
                "\r\n\r\n────────────────────────────────\r\n" +
                "Сторонние компоненты и лицензии\r\n" +
                "────────────────────────────────\r\n\r\n" +
                Properties.Resources.THIRD_PARTY;
            this.LicenseTextBox.SelectionStart = 0;
            this.LicenseTextBox.SelectionLength = 0;
        }

        private void LoadNeZnaikaLogo()
        {
            try
            {
                Assembly assembly = typeof(AboutBox).Assembly;
                using (System.IO.Stream stream = assembly.GetManifestResourceStream(OwlResourceName))
                {
                    if (stream == null)
                        return;

                    // Clone the image so the PictureBox does not depend on the lifetime
                    // of the manifest-resource stream after this method returns.
                    using (Image source = Image.FromStream(stream, true, true))
                        this.logoPictureBox.Image = new Bitmap(source);
                }
            }
            catch
            {
                // The About dialog must remain usable even if Windows cannot decode the
                // optional logo on a particular machine.
                this.logoPictureBox.Image = null;
            }
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
