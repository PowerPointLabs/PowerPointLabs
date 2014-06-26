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
            this.components = new System.ComponentModel.Container();
            this.motherTableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.myShapePanel = new Stepi.UI.ExtendedPanel();
            this.myShapeFlowLayout = new System.Windows.Forms.FlowLayoutPanel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.panel5 = new System.Windows.Forms.Panel();
            this.panel7 = new System.Windows.Forms.Panel();
            this.contextMenuStrip = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.removeShapeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.motherTableLayoutPanel.SuspendLayout();
            this.myShapePanel.SuspendLayout();
            this.contextMenuStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // motherTableLayoutPanel
            // 
            this.motherTableLayoutPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.motherTableLayoutPanel.AutoScroll = true;
            this.motherTableLayoutPanel.ColumnCount = 1;
            this.motherTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.motherTableLayoutPanel.Controls.Add(this.myShapePanel, 0, 0);
            this.motherTableLayoutPanel.Location = new System.Drawing.Point(3, 0);
            this.motherTableLayoutPanel.Margin = new System.Windows.Forms.Padding(0);
            this.motherTableLayoutPanel.MaximumSize = new System.Drawing.Size(500, 336);
            this.motherTableLayoutPanel.Name = "motherTableLayoutPanel";
            this.motherTableLayoutPanel.RowCount = 2;
            this.motherTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.motherTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 1F));
            this.motherTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.motherTableLayoutPanel.Size = new System.Drawing.Size(250, 221);
            this.motherTableLayoutPanel.TabIndex = 0;
            // 
            // myShapePanel
            // 
            this.myShapePanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.myShapePanel.BorderColor = System.Drawing.Color.Gray;
            this.myShapePanel.CaptionColorOne = System.Drawing.Color.White;
            this.myShapePanel.CaptionColorTwo = System.Drawing.Color.FromArgb(((int)(((byte)(155)))), ((int)(((byte)(255)))), ((int)(((byte)(165)))), ((int)(((byte)(0)))));
            this.myShapePanel.CaptionFont = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.myShapePanel.CaptionSize = 20;
            this.myShapePanel.CaptionText = "My Shape";
            this.myShapePanel.CaptionTextColor = System.Drawing.Color.Black;
            this.myShapePanel.Controls.Add(this.myShapeFlowLayout);
            this.myShapePanel.DirectionCtrlColor = System.Drawing.Color.DarkGray;
            this.myShapePanel.DirectionCtrlHoverColor = System.Drawing.Color.Orange;
            this.myShapePanel.Location = new System.Drawing.Point(0, 0);
            this.myShapePanel.Margin = new System.Windows.Forms.Padding(0);
            this.myShapePanel.Name = "myShapePanel";
            this.myShapePanel.Size = new System.Drawing.Size(250, 213);
            this.myShapePanel.State = Stepi.UI.ExtendedPanelState.Collapsed;
            this.myShapePanel.TabIndex = 4;
            // 
            // myShapeFlowLayout
            // 
            this.myShapeFlowLayout.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.myShapeFlowLayout.Location = new System.Drawing.Point(3, 22);
            this.myShapeFlowLayout.Name = "myShapeFlowLayout";
            this.myShapeFlowLayout.Size = new System.Drawing.Size(244, 188);
            this.myShapeFlowLayout.TabIndex = 1;
            // 
            // panel4
            // 
            this.panel4.BackgroundImage = global::PowerPointLabs.Properties.Resources.Pause;
            this.panel4.Location = new System.Drawing.Point(3, 3);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(32, 32);
            this.panel4.TabIndex = 3;
            // 
            // panel5
            // 
            this.panel5.BackgroundImage = global::PowerPointLabs.Properties.Resources.Record;
            this.panel5.Location = new System.Drawing.Point(41, 3);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(32, 32);
            this.panel5.TabIndex = 4;
            // 
            // panel7
            // 
            this.panel7.BackgroundImage = global::PowerPointLabs.Properties.Resources.Stop;
            this.panel7.Location = new System.Drawing.Point(79, 3);
            this.panel7.Name = "panel7";
            this.panel7.Size = new System.Drawing.Size(32, 32);
            this.panel7.TabIndex = 5;
            // 
            // contextMenuStrip
            // 
            this.contextMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.removeShapeToolStripMenuItem});
            this.contextMenuStrip.Name = "contextMenuStrip";
            this.contextMenuStrip.Size = new System.Drawing.Size(164, 26);
            this.contextMenuStrip.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.ContextMenuStripItemClicked);
            // 
            // removeShapeToolStripMenuItem
            // 
            this.removeShapeToolStripMenuItem.Name = "removeShapeToolStripMenuItem";
            this.removeShapeToolStripMenuItem.Size = new System.Drawing.Size(163, 22);
            this.removeShapeToolStripMenuItem.Text = "Remove Shape";
            // 
            // CustomShapePane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.motherTableLayoutPanel);
            this.Name = "CustomShapePane";
            this.Size = new System.Drawing.Size(257, 459);
            this.motherTableLayoutPanel.ResumeLayout(false);
            this.myShapePanel.ResumeLayout(false);
            this.contextMenuStrip.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel motherTableLayoutPanel;
        private ExtendedPanel myShapePanel;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Panel panel7;
        private System.Windows.Forms.FlowLayoutPanel myShapeFlowLayout;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip;
        private System.Windows.Forms.ToolStripMenuItem removeShapeToolStripMenuItem;

    }
}
