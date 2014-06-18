using System.Drawing;
using Stepi.UI;

namespace PowerPointLabs
{
    partial class CustomShapePane
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
            this.extendedPanel1 = new Stepi.UI.ExtendedPanel();
            this.extendedPanel2 = new Stepi.UI.ExtendedPanel();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            this.extendedPanel1.SuspendLayout();
            this.flowLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // extendedPanel1
            // 
            this.extendedPanel1.BorderColor = System.Drawing.Color.Gray;
            this.extendedPanel1.CaptionColorOne = System.Drawing.SystemColors.ControlLight;
            this.extendedPanel1.CaptionColorTwo = System.Drawing.SystemColors.ControlDark;
            this.extendedPanel1.CaptionFont = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.extendedPanel1.CaptionSize = 20;
            this.extendedPanel1.CaptionText = "My Shape";
            this.extendedPanel1.CaptionTextColor = System.Drawing.Color.Black;
            this.extendedPanel1.Controls.Add(this.flowLayoutPanel2);
            this.extendedPanel1.DirectionCtrlColor = System.Drawing.Color.AntiqueWhite;
            this.extendedPanel1.DirectionCtrlHoverColor = System.Drawing.Color.Aqua;
            this.extendedPanel1.Location = new System.Drawing.Point(3, 3);
            this.extendedPanel1.Name = "extendedPanel1";
            this.extendedPanel1.Size = new System.Drawing.Size(249, 134);
            this.extendedPanel1.TabIndex = 0;
            // 
            // extendedPanel2
            // 
            this.extendedPanel2.BorderColor = System.Drawing.Color.Gray;
            this.extendedPanel2.CaptionColorOne = System.Drawing.Color.White;
            this.extendedPanel2.CaptionColorTwo = System.Drawing.Color.FromArgb(((int)(((byte)(155)))), ((int)(((byte)(255)))), ((int)(((byte)(165)))), ((int)(((byte)(0)))));
            this.extendedPanel2.CaptionFont = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.extendedPanel2.CaptionSize = 20;
            this.extendedPanel2.CaptionTextColor = System.Drawing.Color.Black;
            this.extendedPanel2.DirectionCtrlColor = System.Drawing.Color.DarkGray;
            this.extendedPanel2.DirectionCtrlHoverColor = System.Drawing.Color.Orange;
            this.extendedPanel2.Location = new System.Drawing.Point(3, 143);
            this.extendedPanel2.Name = "extendedPanel2";
            this.extendedPanel2.Size = new System.Drawing.Size(249, 89);
            this.extendedPanel2.TabIndex = 1;
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Controls.Add(this.extendedPanel1);
            this.flowLayoutPanel1.Controls.Add(this.extendedPanel2);
            this.flowLayoutPanel1.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(3, 3);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(260, 354);
            this.flowLayoutPanel1.TabIndex = 2;
            // 
            // flowLayoutPanel2
            // 
            this.flowLayoutPanel2.AllowDrop = true;
            this.flowLayoutPanel2.Location = new System.Drawing.Point(3, 23);
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            this.flowLayoutPanel2.Size = new System.Drawing.Size(243, 108);
            this.flowLayoutPanel2.TabIndex = 1;
            // 
            // CustomShapePane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.flowLayoutPanel1);
            this.Name = "CustomShapePane";
            this.Size = new System.Drawing.Size(266, 360);
            this.extendedPanel1.ResumeLayout(false);
            this.flowLayoutPanel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private ExtendedPanel extendedPanel1;
        private ExtendedPanel extendedPanel2;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel2;
    }
}
