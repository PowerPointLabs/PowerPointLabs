namespace PowerPointLabs.Views
{
    partial class AutoNarrateDialogBox
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
            this.allSlides = new System.Windows.Forms.CheckBox();
            this.preview = new System.Windows.Forms.CheckBox();
            this.cancel = new System.Windows.Forms.Button();
            this.ok = new System.Windows.Forms.Button();
            this.defaultVoice = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // allSlides
            // 
            this.allSlides.AutoSize = true;
            this.allSlides.Location = new System.Drawing.Point(24, 49);
            this.allSlides.Margin = new System.Windows.Forms.Padding(4);
            this.allSlides.Name = "allSlides";
            this.allSlides.Size = new System.Drawing.Size(87, 21);
            this.allSlides.TabIndex = 0;
            this.allSlides.Text = "All Slides";
            this.allSlides.UseVisualStyleBackColor = true;
            // 
            // preview
            // 
            this.preview.AutoSize = true;
            this.preview.Location = new System.Drawing.Point(24, 78);
            this.preview.Margin = new System.Windows.Forms.Padding(4);
            this.preview.Name = "preview";
            this.preview.Size = new System.Drawing.Size(79, 21);
            this.preview.TabIndex = 1;
            this.preview.Text = "Preview";
            this.preview.UseVisualStyleBackColor = true;
            // 
            // cancel
            // 
            this.cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancel.Location = new System.Drawing.Point(295, 98);
            this.cancel.Margin = new System.Windows.Forms.Padding(4);
            this.cancel.Name = "cancel";
            this.cancel.Size = new System.Drawing.Size(100, 28);
            this.cancel.TabIndex = 2;
            this.cancel.Text = "Cancel";
            this.cancel.UseVisualStyleBackColor = true;
            this.cancel.Click += new System.EventHandler(this.cancel_Click);
            // 
            // ok
            // 
            this.ok.Location = new System.Drawing.Point(187, 98);
            this.ok.Margin = new System.Windows.Forms.Padding(4);
            this.ok.Name = "ok";
            this.ok.Size = new System.Drawing.Size(100, 28);
            this.ok.TabIndex = 3;
            this.ok.Text = "OK";
            this.ok.UseVisualStyleBackColor = true;
            this.ok.Click += new System.EventHandler(this.ok_Click);
            // 
            // defaultVoice
            // 
            this.defaultVoice.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.defaultVoice.FormattingEnabled = true;
            this.defaultVoice.Location = new System.Drawing.Point(130, 16);
            this.defaultVoice.Margin = new System.Windows.Forms.Padding(4);
            this.defaultVoice.Name = "defaultVoice";
            this.defaultVoice.Size = new System.Drawing.Size(268, 24);
            this.defaultVoice.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(21, 19);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(92, 17);
            this.label1.TabIndex = 5;
            this.label1.Text = "Default Voice";
            // 
            // AutoNarrateDialogBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.cancel;
            this.ClientSize = new System.Drawing.Size(411, 137);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.defaultVoice);
            this.Controls.Add(this.ok);
            this.Controls.Add(this.cancel);
            this.Controls.Add(this.preview);
            this.Controls.Add(this.allSlides);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AutoNarrateDialogBox";
            this.ShowIcon = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Auto Narrate Properties";
            this.Load += new System.EventHandler(this.AutoNarrateDialogBox_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox allSlides;
        private System.Windows.Forms.CheckBox preview;
        private System.Windows.Forms.Button cancel;
        private System.Windows.Forms.Button ok;
        private System.Windows.Forms.ComboBox defaultVoice;
        private System.Windows.Forms.Label label1;
    }
}