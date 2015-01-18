namespace PowerPointLabs.Views
{
    partial class BulletAgendaSettingsDialog
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
            this.highLightLabel = new System.Windows.Forms.Label();
            this.dimColorLabel = new System.Windows.Forms.Label();
            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.dimColorBox = new System.Windows.Forms.PictureBox();
            this.higlightColorBox = new System.Windows.Forms.PictureBox();
            this.defaultColorLabel = new System.Windows.Forms.Label();
            this.defaultColorBox = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.dimColorBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.higlightColorBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.defaultColorBox)).BeginInit();
            this.SuspendLayout();
            // 
            // highLightLabel
            // 
            this.highLightLabel.AutoSize = true;
            this.highLightLabel.Location = new System.Drawing.Point(12, 17);
            this.highLightLabel.Name = "highLightLabel";
            this.highLightLabel.Size = new System.Drawing.Size(95, 12);
            this.highLightLabel.TabIndex = 0;
            this.highLightLabel.Text = "Highlight Color";
            // 
            // dimColorLabel
            // 
            this.dimColorLabel.AutoSize = true;
            this.dimColorLabel.Location = new System.Drawing.Point(12, 49);
            this.dimColorLabel.Name = "dimColorLabel";
            this.dimColorLabel.Size = new System.Drawing.Size(59, 12);
            this.dimColorLabel.TabIndex = 1;
            this.dimColorLabel.Text = "Dim Color";
            // 
            // okButton
            // 
            this.okButton.Location = new System.Drawing.Point(86, 121);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(75, 23);
            this.okButton.TabIndex = 2;
            this.okButton.Text = "Ok";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.OkButtonClick);
            // 
            // cancelButton
            // 
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Location = new System.Drawing.Point(167, 121);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 23);
            this.cancelButton.TabIndex = 3;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            // 
            // dimColorBox
            // 
            this.dimColorBox.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dimColorBox.Location = new System.Drawing.Point(204, 47);
            this.dimColorBox.Name = "dimColorBox";
            this.dimColorBox.Size = new System.Drawing.Size(38, 19);
            this.dimColorBox.TabIndex = 5;
            this.dimColorBox.TabStop = false;
            this.dimColorBox.Click += new System.EventHandler(this.DimColorBoxClick);
            // 
            // higlightColorBox
            // 
            this.higlightColorBox.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.higlightColorBox.ErrorImage = null;
            this.higlightColorBox.InitialImage = null;
            this.higlightColorBox.Location = new System.Drawing.Point(204, 16);
            this.higlightColorBox.Name = "higlightColorBox";
            this.higlightColorBox.Size = new System.Drawing.Size(38, 19);
            this.higlightColorBox.TabIndex = 4;
            this.higlightColorBox.TabStop = false;
            this.higlightColorBox.Click += new System.EventHandler(this.HiglightColorBoxClick);
            // 
            // defaultColorLabel
            // 
            this.defaultColorLabel.AutoSize = true;
            this.defaultColorLabel.Location = new System.Drawing.Point(12, 82);
            this.defaultColorLabel.Name = "defaultColorLabel";
            this.defaultColorLabel.Size = new System.Drawing.Size(83, 12);
            this.defaultColorLabel.TabIndex = 6;
            this.defaultColorLabel.Text = "Default Color";
            // 
            // defaultColorBox
            // 
            this.defaultColorBox.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.defaultColorBox.ErrorImage = null;
            this.defaultColorBox.InitialImage = null;
            this.defaultColorBox.Location = new System.Drawing.Point(204, 79);
            this.defaultColorBox.Name = "defaultColorBox";
            this.defaultColorBox.Size = new System.Drawing.Size(38, 19);
            this.defaultColorBox.TabIndex = 7;
            this.defaultColorBox.TabStop = false;
            this.defaultColorBox.Click += new System.EventHandler(this.DefaultColorBoxClick);
            // 
            // BulletAgendaSettingsDialog
            // 
            this.AcceptButton = this.okButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.cancelButton;
            this.ClientSize = new System.Drawing.Size(267, 156);
            this.Controls.Add(this.defaultColorBox);
            this.Controls.Add(this.defaultColorLabel);
            this.Controls.Add(this.dimColorBox);
            this.Controls.Add(this.higlightColorBox);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.dimColorLabel);
            this.Controls.Add(this.highLightLabel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "BulletAgendaSettingsDialog";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Bullet Agenda Settings";
            ((System.ComponentModel.ISupportInitialize)(this.dimColorBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.higlightColorBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.defaultColorBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label highLightLabel;
        private System.Windows.Forms.Label dimColorLabel;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.PictureBox higlightColorBox;
        private System.Windows.Forms.PictureBox dimColorBox;
        private System.Windows.Forms.Label defaultColorLabel;
        private System.Windows.Forms.PictureBox defaultColorBox;
    }
}