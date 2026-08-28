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

        private const string NeZnaikaAboutText =
            "Надстройка, созданная во имя текста, правок и человеческих страданий.\r\n\r\n" +
            "НеZнайка умеет ковырять текст, шаманить над формулировками, совершать обряды форматирования и делать прочие вещи, которые нормальный человек предпочёл бы поручить кому-нибудь другому.\r\n\r\n" +
            "Если документ стал лучше — так и было задумано.\r\n\r\n" +
            "Если стал хуже — это авторский стиль.\r\n\r\n" +
            "Если кончилась оперативка — значит, таинство началось.\r\n\r\n" +
            "Вместе с НеZнайкой мы натянем любую сову на глобус!\r\n\r\n" +
            "И не забывайте страдать!";

        private const string NeZnaikaLogoBase64 =
            "iVBORw0KGgoAAAANSUhEUgAAAMUAAADbCAMAAADqOqUpAAADAFBMVEX6+vrq6uoBsP76yG/cBhL704wArvzgBhP3vFT5wFzY2NfGxsYIBQW4uLj6zoEpuvqv5f1Mxv0bGxvW8f5y0v786srG7P1WVlaL2v6np6f73KZHR0f019k5wf784rRoZ2cnJib98ds4NjPwt7voummWlZXqhowyAgXmZ27SuG/kR1CzBQ/iJjHaFyPaJjFjzf2qiExMAgbzx8niFyPmWmLtpqpwAwp6enqIh4fslJkmk9SRuJpSRzHqe4KUBA25lVXkOUNjZZfJp2qdOFi0mGXoc3olIx53ZkbbER3vnaLmU1kJltZrWTjhERzvvsEJisZEOyqMdUyb4P7DnFcENExOtMjPpVrjMjkFea0jGxpyta/ZRU45s9h5ZD3cND7cO0QAKDoFR2UEZpUrquhIve+st4es2vD23+ADgbo2hLo1isg3ptSBbUiEtqKMus7eU1sEGSUbRlk2WHRbUDtRaotRe45GfLVctL1+TXR+UXtyWYdxqMu9JTqiOFizQk29uXvNv8DAwL/lIx7rVjToZDj0nEoanOQqTl4mRmYxZHs7boUxd69OLTBSNlJZbaF6LTJ0Lkd5N1VjQkRlWEBoiblnuuCGGR+JSW+cbnOYqJCBq76bz+agGCCiUG+laYe6g4apjqqqkpOtsMqnx7Khwc+xzue108bfWmLBY2jSglbIjI/NnGXesF/bxoTpRirweD3xhEL1r1Hhv4EAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABEtkB4AAABAHRSTlP//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAPlDdoQAADbBJREFUeNrt3Yd/2lYeAHA9EAbEFMtgxB7BmJFgY0Nixz7XO7ZjO46zmrRpm6S713l33b22t/fee++7v/Dek8Ag8QSSkADz0e/z6ScYjKwvv/d7S3JNAFwQZy10ha7QFbpCVwxDQYyDgtAVukJX6ApdoSt0ha7QFbpCV+gKXaErdIWu0BW6QleMnCLRiLOqSPj9/igbUyjgv/CJM6ZI+KNTpmY4nU4jDJMp6k+cHQVMwlTLcMpAj2BS+pMMTOGPmoTRyAabEegYeQWvKeEZfTkGosDkoYOB8qG4WQ1Agc9Dk9FSGJ3QMaoK0UR0MoxKGZor/FOmrsFTIEdi9BSJXghBMiAjmhg5RdRkkqlQlA1NFYnolARFB0N+l6uponcmcMkwOqdGSCElE1iF/GxoqPCjM5TCMHbG1KgoYO/kVKww+UdDkYAIozRF/wytFIko+oyVK+S1Ka0UsD0ZpSpwCHnJ0EiRYBH9KGQxNFL4TcZ+FcbosBVce5KmcIopZCRDG0XUZJSqcIoorFan9AmVJoqE6XQy4VSosFqtRqd/qAq/LIVRTCG9MrRQ+KOqKKxWyZWhiaKt2+EznP9bXVtbW512ChHp1WUY16YFiugQFa1UCBnp1cPXX//+i2tpYVVMn7+7u7h7dzndprBYjFLrWwvFFG+x0MZwrm0uJpPJxc21RjZOEYe7B8nkweKLy/9opQIqpDYpDRQJkxHPcDoP4xseGMnDVSevf7q2GL8Fn7+1ePcaT2EcnsJvMuIZzunNjVusYnPNyOtklxdZnedgl6ewSu2lBqA4ZaRXN9lzbSjaRorlg8YLu8ut2rbAZAxNkYiajJ0MBDFOtymc7cOdMBcNhEXq/Fx9RXMO1eEwGdOH8Tg8143Fw1Ujb5xY3T2AbW0DvnDSjrBYJXZSA1JwDiPsozY24geba2n+i+nlu8kNT/zg8MQyMgqxSarTOb324uLi4ub5tPCl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";

        public AboutBox()
        {
            try
            {
                InitializeComponent();
                this.Text = "О программе — НеZнайка";
                this.labelProductName.Text = "НеZнайка";
                this.labelVersion.Text = "Версия " + AssemblyVersion;
                this.labelCopyright.Text = AssemblyCopyright;
                this.labelCompanyName.Text = "Локальная AI-надстройка для Microsoft Word";

                ApplyNeZnaikaLogo();

                this.LicenseTextBox.Multiline = true;
                this.LicenseTextBox.ScrollBars = ScrollBars.Vertical;
                this.LicenseTextBox.Text =
                    NeZnaikaAboutText +
                    "\r\n\r\n────────────────────────────────\r\n" +
                    "Сторонние компоненты и лицензии\r\n────────────────────────────────\r\n\r\n" +
                    Properties.Resources.THIRD_PARTY;
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void ApplyNeZnaikaLogo()
        {
            byte[] imageBytes = Convert.FromBase64String(NeZnaikaLogoBase64);
            using (MemoryStream stream = new MemoryStream(imageBytes))
            using (Image sourceImage = Image.FromStream(stream))
            {
                this.logoPictureBox.Image = new Bitmap(sourceImage);
            }

            this.logoPictureBox.SizeMode = PictureBoxSizeMode.Zoom;
            this.logoPictureBox.BackColor = Color.White;
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
