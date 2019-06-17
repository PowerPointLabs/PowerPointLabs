using System.Windows.Controls;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ResizeLab
{
    partial class ResizeLabPane: IWpfContainer
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

            WPF = new System.Windows.Forms.Integration.ElementHost();
            resizePaneWPF = new ResizeLabPaneWPF();
            SuspendLayout();
            // 
            // WPF
            // 
            WPF.Dock = System.Windows.Forms.DockStyle.Fill;
            WPF.ForeColor = System.Drawing.SystemColors.ButtonFace;
            WPF.Location = new System.Drawing.Point(0, 0);
            WPF.Name = "WPF";
            WPF.Size = new System.Drawing.Size(304, 883);
            WPF.Text = "WPF";
            WPF.Child = resizePaneWPF;
            // 
            // ResizePane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.WPF);
            this.Name = "ResizePane";
            this.Size = new System.Drawing.Size(300, 883);
            this.ResumeLayout(false);

        }
        public ResizeLabPaneWPF resizePaneWPF { get; private set; }

        private System.Windows.Forms.Integration.ElementHost WPF;
        #endregion

        public Control WpfControl => resizePaneWPF;
    }
}
