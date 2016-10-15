namespace abacode_senior_project
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.logo = new System.Windows.Forms.PictureBox();
            this.pathToReportLabel = new System.Windows.Forms.Label();
            this.pathTextBox = new System.Windows.Forms.TextBox();
            this.typeOfReportLabel = new System.Windows.Forms.Label();
            this.openVASRadio = new System.Windows.Forms.RadioButton();
            this.NessusRadio = new System.Windows.Forms.RadioButton();
            this.browseButton = new System.Windows.Forms.Button();
            this.startButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.logo)).BeginInit();
            this.SuspendLayout();
            // 
            // logo
            // 
            this.logo.Dock = System.Windows.Forms.DockStyle.Top;
            this.logo.Image = ((System.Drawing.Image)(resources.GetObject("logo.Image")));
            this.logo.Location = new System.Drawing.Point(0, 0);
            this.logo.Name = "logo";
            this.logo.Size = new System.Drawing.Size(456, 142);
            this.logo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.logo.TabIndex = 0;
            this.logo.TabStop = false;
            // 
            // pathToReportLabel
            // 
            this.pathToReportLabel.AutoSize = true;
            this.pathToReportLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.pathToReportLabel.Location = new System.Drawing.Point(12, 160);
            this.pathToReportLabel.Name = "pathToReportLabel";
            this.pathToReportLabel.Size = new System.Drawing.Size(125, 20);
            this.pathToReportLabel.TabIndex = 2;
            this.pathToReportLabel.Text = "Path to report:";
            // 
            // pathTextBox
            // 
            this.pathTextBox.Location = new System.Drawing.Point(143, 162);
            this.pathTextBox.Name = "pathTextBox";
            this.pathTextBox.Size = new System.Drawing.Size(216, 20);
            this.pathTextBox.TabIndex = 3;
            // 
            // typeOfReportLabel
            // 
            this.typeOfReportLabel.AutoSize = true;
            this.typeOfReportLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.typeOfReportLabel.Location = new System.Drawing.Point(12, 208);
            this.typeOfReportLabel.Name = "typeOfReportLabel";
            this.typeOfReportLabel.Size = new System.Drawing.Size(126, 20);
            this.typeOfReportLabel.TabIndex = 4;
            this.typeOfReportLabel.Text = "Type of report:";
            // 
            // openVASRadio
            // 
            this.openVASRadio.AutoSize = true;
            this.openVASRadio.Checked = true;
            this.openVASRadio.Location = new System.Drawing.Point(145, 210);
            this.openVASRadio.Name = "openVASRadio";
            this.openVASRadio.Size = new System.Drawing.Size(70, 17);
            this.openVASRadio.TabIndex = 5;
            this.openVASRadio.TabStop = true;
            this.openVASRadio.Text = "openVAS";
            this.openVASRadio.UseVisualStyleBackColor = true;
            // 
            // NessusRadio
            // 
            this.NessusRadio.AutoSize = true;
            this.NessusRadio.Location = new System.Drawing.Point(222, 210);
            this.NessusRadio.Name = "NessusRadio";
            this.NessusRadio.Size = new System.Drawing.Size(60, 17);
            this.NessusRadio.TabIndex = 6;
            this.NessusRadio.TabStop = true;
            this.NessusRadio.Text = "Nessus";
            this.NessusRadio.UseVisualStyleBackColor = true;
            // 
            // browseButton
            // 
            this.browseButton.Location = new System.Drawing.Point(365, 160);
            this.browseButton.Name = "browseButton";
            this.browseButton.Size = new System.Drawing.Size(75, 23);
            this.browseButton.TabIndex = 7;
            this.browseButton.Text = "Browse";
            this.browseButton.UseVisualStyleBackColor = true;
            this.browseButton.Click += new System.EventHandler(this.browseButton_Click);
            // 
            // startButton
            // 
            this.startButton.ForeColor = System.Drawing.SystemColors.ControlText;
            this.startButton.Location = new System.Drawing.Point(193, 256);
            this.startButton.Name = "startButton";
            this.startButton.Size = new System.Drawing.Size(75, 23);
            this.startButton.TabIndex = 8;
            this.startButton.Text = "Start";
            this.startButton.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(456, 291);
            this.Controls.Add(this.startButton);
            this.Controls.Add(this.browseButton);
            this.Controls.Add(this.NessusRadio);
            this.Controls.Add(this.openVASRadio);
            this.Controls.Add(this.typeOfReportLabel);
            this.Controls.Add(this.pathTextBox);
            this.Controls.Add(this.pathToReportLabel);
            this.Controls.Add(this.logo);
            this.Name = "Form1";
            this.Text = "Abacode Senior Project";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.logo)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox logo;
        private System.Windows.Forms.Label pathToReportLabel;
        private System.Windows.Forms.TextBox pathTextBox;
        private System.Windows.Forms.Label typeOfReportLabel;
        private System.Windows.Forms.RadioButton openVASRadio;
        private System.Windows.Forms.RadioButton NessusRadio;
        private System.Windows.Forms.Button browseButton;
        private System.Windows.Forms.Button startButton;
    }
}

