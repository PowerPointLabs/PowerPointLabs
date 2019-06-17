using System.Windows.Controls;
using PowerPointLabs.Utils;

namespace PowerPointLabs.TimerLab
{
    partial class TimerPane: IWpfContainer
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
            this.wpf = new System.Windows.Forms.Integration.ElementHost();
            this.TimerPaneWPF = new PowerPointLabs.TimerLab.TimerLabPaneWPF();
            this.SuspendLayout();
            // 
            // wpf
            // 
            this.wpf.Dock = System.Windows.Forms.DockStyle.Fill;
            this.wpf.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.wpf.Location = new System.Drawing.Point(0, 0);
            this.wpf.Name = "wpf";
            this.wpf.Size = new System.Drawing.Size(300, 883);
            this.wpf.TabIndex = 0;
            this.wpf.Text = "wpf";
            this.wpf.Child = this.TimerPaneWPF;
            // 
            // TimerPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.wpf);
            this.Name = "TimerPane";
            this.Size = new System.Drawing.Size(300, 883);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Integration.ElementHost wpf;
        public TimerLabPaneWPF TimerPaneWPF { get; private set; }

        public Control WpfControl
        {
            get
            {
                return TimerPaneWPF;
            }
        }
    }
}
