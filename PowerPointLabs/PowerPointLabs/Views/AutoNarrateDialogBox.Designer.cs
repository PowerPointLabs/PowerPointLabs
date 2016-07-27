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
            this.preview = new System.Windows.Forms.CheckBox();
            this.cancel = new System.Windows.Forms.Button();
            this.ok = new System.Windows.Forms.Button();
            this.defaultVoice = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // preview
            // 
            this.preview.AutoSize = true;
            this.preview.Location = new System.Drawing.Point(19, 55);
            this.preview.Name = "preview";
            this.preview.Size = new System.Drawing.Size(64, 17);
            this.preview.TabIndex = 1;
            this.preview.Text = "Preview";
            this.preview.UseVisualStyleBackColor = true;
            // 
            // cancel
            // 
            this.cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancel.Location = new System.Drawing.Point(221, 80);
            this.cancel.Name = "cancel";
            this.cancel.Size = new System.Drawing.Size(75, 23);
            this.cancel.TabIndex = 2;
            this.cancel.Text = "Cancel";
            this.cancel.UseVisualStyleBackColor = true;
            this.cancel.Click += new System.EventHandler(this.Cancel_Click);
            // 
            // ok
            // 
            this.ok.Location = new System.Drawing.Point(140, 80);
            this.ok.Name = "ok";
            this.ok.Size = new System.Drawing.Size(75, 23);
            this.ok.TabIndex = 3;
            this.ok.Text = "OK";
            this.ok.UseVisualStyleBackColor = true;
            this.ok.Click += new System.EventHandler(this.Ok_Click);
            // 
            // defaultVoice
            // 
            this.defaultVoice.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.defaultVoice.FormattingEnabled = true;
            this.defaultVoice.Location = new System.Drawing.Point(98, 13);
            this.defaultVoice.Name = "defaultVoice";
            this.defaultVoice.Size = new System.Drawing.Size(202, 21);
            this.defaultVoice.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Default Voice";
            // 
            // AutoNarrateDialogBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.cancel;
            this.ClientSize = new System.Drawing.Size(308, 111);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.defaultVoice);
            this.Controls.Add(this.ok);
            this.Controls.Add(this.cancel);
            this.Controls.Add(this.preview);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AutoNarrateDialogBox";
            this.ShowIcon = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Narrations Lab Properties";
            this.Load += new System.EventHandler(this.AutoNarrateDialogBox_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox preview;
        private System.Windows.Forms.Button cancel;
        private System.Windows.Forms.Button ok;
        private System.Windows.Forms.ComboBox defaultVoice;
        private System.Windows.Forms.Label label1;
    }
}