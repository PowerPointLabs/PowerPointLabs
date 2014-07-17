namespace PowerPointLabs.ColorPicker
{
    partial class ColorInformationDialog
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
            this.components = new System.ComponentModel.Container();
            this.selectedColorPanel = new System.Windows.Forms.Panel();
            this.hexTextBox = new System.Windows.Forms.TextBox();
            this.rgbTextBox = new System.Windows.Forms.TextBox();
            this.HSLTextBox = new System.Windows.Forms.TextBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // selectedColorPanel
            // 
            this.selectedColorPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.selectedColorPanel.Location = new System.Drawing.Point(12, 12);
            this.selectedColorPanel.Name = "selectedColorPanel";
            this.selectedColorPanel.Size = new System.Drawing.Size(118, 68);
            this.selectedColorPanel.TabIndex = 0;
            // 
            // hexTextBox
            // 
            this.hexTextBox.BackColor = System.Drawing.SystemColors.ControlLight;
            this.hexTextBox.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.hexTextBox.HideSelection = false;
            this.hexTextBox.Location = new System.Drawing.Point(12, 86);
            this.hexTextBox.Name = "hexTextBox";
            this.hexTextBox.ReadOnly = true;
            this.hexTextBox.Size = new System.Drawing.Size(118, 27);
            this.hexTextBox.TabIndex = 3;
            // 
            // rgbTextBox
            // 
            this.rgbTextBox.BackColor = System.Drawing.SystemColors.ControlLight;
            this.rgbTextBox.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rgbTextBox.HideSelection = false;
            this.rgbTextBox.Location = new System.Drawing.Point(12, 118);
            this.rgbTextBox.Name = "rgbTextBox";
            this.rgbTextBox.ReadOnly = true;
            this.rgbTextBox.Size = new System.Drawing.Size(118, 27);
            this.rgbTextBox.TabIndex = 4;
            // 
            // HSLTextBox
            // 
            this.HSLTextBox.BackColor = System.Drawing.SystemColors.ControlLight;
            this.HSLTextBox.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.HSLTextBox.HideSelection = false;
            this.HSLTextBox.Location = new System.Drawing.Point(12, 150);
            this.HSLTextBox.Name = "HSLTextBox";
            this.HSLTextBox.ReadOnly = true;
            this.HSLTextBox.Size = new System.Drawing.Size(118, 27);
            this.HSLTextBox.TabIndex = 5;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Location = new System.Drawing.Point(0, 186);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(142, 22);
            this.statusStrip1.SizingGrip = false;
            this.statusStrip1.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(4, 190);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 13);
            this.label1.TabIndex = 7;
            // 
            // ColorInformationDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(142, 208);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.HSLTextBox);
            this.Controls.Add(this.rgbTextBox);
            this.Controls.Add(this.hexTextBox);
            this.Controls.Add(this.selectedColorPanel);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(158, 246);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(158, 246);
            this.Name = "ColorInformationDialog";
            this.ShowIcon = false;
            this.Text = "Color Info";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel selectedColorPanel;
        private System.Windows.Forms.TextBox hexTextBox;
        private System.Windows.Forms.TextBox rgbTextBox;
        private System.Windows.Forms.TextBox HSLTextBox;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.Label label1;
    }
}