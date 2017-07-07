namespace PowerPointLabs.ShapesLab
{
    partial class LabeledThumbnail
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.motherPanel = new PowerPointLabs.BufferedPanel();
            this.labelTextBox = new System.Windows.Forms.TextBox();
            this.thumbnailPanel = new PowerPointLabs.BufferedPanel();
            this.nameLabelToolTip = new System.Windows.Forms.ToolTip(this.components);
            this.motherPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // motherPanel
            // 
            this.motherPanel.Controls.Add(this.labelTextBox);
            this.motherPanel.Controls.Add(this.thumbnailPanel);
            this.motherPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.motherPanel.Location = new System.Drawing.Point(0, 0);
            this.motherPanel.Name = "motherPanel";
            this.motherPanel.Size = new System.Drawing.Size(118, 50);
            this.motherPanel.TabIndex = 2;
            // 
            // labelTextBox
            // 
            this.labelTextBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.labelTextBox.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.labelTextBox.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelTextBox.Location = new System.Drawing.Point(52, 1);
            this.labelTextBox.Margin = new System.Windows.Forms.Padding(0, 3, 0, 3);
            this.labelTextBox.Multiline = true;
            this.labelTextBox.Name = "labelTextBox";
            this.labelTextBox.Size = new System.Drawing.Size(65, 48);
            this.labelTextBox.TabIndex = 1;
            // 
            // thumbnailPanel
            // 
            //this.thumbnailPanel.BackColor = System.Drawing.Color.Transparent;
            this.thumbnailPanel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.thumbnailPanel.Location = new System.Drawing.Point(0, 0);
            this.thumbnailPanel.Margin = new System.Windows.Forms.Padding(0, 3, 0, 3);
            this.thumbnailPanel.Name = "thumbnailPanel";
            this.thumbnailPanel.Size = new System.Drawing.Size(50, 50);
            this.thumbnailPanel.TabIndex = 0;
            // 
            // nameLabelToolTip
            // 
            this.nameLabelToolTip.ShowAlways = true;
            // 
            // LabeledThumbnail
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.motherPanel);
            this.Name = "LabeledThumbnail";
            this.Size = new System.Drawing.Size(118, 50);
            this.motherPanel.ResumeLayout(false);
            this.motherPanel.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private PowerPointLabs.BufferedPanel motherPanel;
        private System.Windows.Forms.TextBox labelTextBox;
        private PowerPointLabs.BufferedPanel thumbnailPanel;
        private System.Windows.Forms.ToolTip nameLabelToolTip;
    }
}
