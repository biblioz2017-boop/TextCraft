namespace TextForge
{
    partial class GenerateUserControl
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GenerateUserControl));
            this.PromptTextBox = new System.Windows.Forms.TextBox();
            this.GenerateButton = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.TemperatureValueLabel = new System.Windows.Forms.Label();
            this.TemperatureLabel = new System.Windows.Forms.Label();
            this.TemperatureTrackBar = new System.Windows.Forms.TrackBar();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TemperatureTrackBar)).BeginInit();
            this.SuspendLayout();

            // PromptTextBox
            resources.ApplyResources(this.PromptTextBox, "PromptTextBox");
            this.PromptTextBox.Name = "PromptTextBox";
            this.PromptTextBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.PromptTextBox_KeyDown);

            // GenerateButton
            resources.ApplyResources(this.GenerateButton, "GenerateButton");
            this.GenerateButton.Name = "GenerateButton";
            this.GenerateButton.Text = "Выполнить";
            this.GenerateButton.UseVisualStyleBackColor = true;
            this.GenerateButton.Click += new System.EventHandler(this.GenerateButton_Click);

            // panel1 (advanced temperature control hidden in simplified UI)
            this.panel1.Controls.Add(this.TemperatureValueLabel);
            this.panel1.Controls.Add(this.TemperatureLabel);
            this.panel1.Controls.Add(this.TemperatureTrackBar);
            resources.ApplyResources(this.panel1, "panel1");
            this.panel1.Name = "panel1";
            this.panel1.Visible = false;

            // TemperatureValueLabel
            resources.ApplyResources(this.TemperatureValueLabel, "TemperatureValueLabel");
            this.TemperatureValueLabel.Name = "TemperatureValueLabel";

            // TemperatureLabel
            resources.ApplyResources(this.TemperatureLabel, "TemperatureLabel");
            this.TemperatureLabel.Name = "TemperatureLabel";

            // TemperatureTrackBar
            resources.ApplyResources(this.TemperatureTrackBar, "TemperatureTrackBar");
            this.TemperatureTrackBar.LargeChange = 1;
            this.TemperatureTrackBar.Name = "TemperatureTrackBar";
            this.TemperatureTrackBar.SmallChange = 1;
            this.TemperatureTrackBar.Value = 1;
            this.TemperatureTrackBar.Scroll += new System.EventHandler(this.TemperatureTrackBar_Scroll);

            // GenerateUserControl
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.GenerateButton);
            this.Controls.Add(this.PromptTextBox);
            this.Name = "GenerateUserControl";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TemperatureTrackBar)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.TextBox PromptTextBox;
        private System.Windows.Forms.Button GenerateButton;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TrackBar TemperatureTrackBar;
        private System.Windows.Forms.Label TemperatureLabel;
        private System.Windows.Forms.Label TemperatureValueLabel;
    }
}
