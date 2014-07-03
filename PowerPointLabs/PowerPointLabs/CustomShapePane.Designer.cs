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
            this.contextMenuStrip = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.removeShapeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.editNameToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.motherTableLayoutPanel.SuspendLayout();
            this.myShapePanel.SuspendLayout();
            this.contextMenuStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // motherTableLayoutPanel
            // 
            this.motherTableLayoutPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.motherTableLayoutPanel.AutoSize = true;
            this.motherTableLayoutPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.motherTableLayoutPanel.ColumnCount = 1;
            this.motherTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.motherTableLayoutPanel.Controls.Add(this.myShapePanel, 0, 0);
            this.motherTableLayoutPanel.Location = new System.Drawing.Point(3, 0);
            this.motherTableLayoutPanel.Margin = new System.Windows.Forms.Padding(0);
            this.motherTableLayoutPanel.MaximumSize = new System.Drawing.Size(500, 500);
            this.motherTableLayoutPanel.MinimumSize = new System.Drawing.Size(120, 50);
            this.motherTableLayoutPanel.Name = "motherTableLayoutPanel";
            this.motherTableLayoutPanel.RowCount = 2;
            this.motherTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.motherTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 1F));
            this.motherTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.motherTableLayoutPanel.Size = new System.Drawing.Size(383, 83);
            this.motherTableLayoutPanel.TabIndex = 0;
            // 
            // myShapePanel
            // 
            this.myShapePanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.myShapePanel.AutoSize = true;
            this.myShapePanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
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
            this.myShapePanel.Size = new System.Drawing.Size(383, 82);
            this.myShapePanel.TabIndex = 4;
            // 
            // myShapeFlowLayout
            // 
            this.myShapeFlowLayout.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.myShapeFlowLayout.AutoScroll = true;
            this.myShapeFlowLayout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.myShapeFlowLayout.Location = new System.Drawing.Point(3, 22);
            this.myShapeFlowLayout.Margin = new System.Windows.Forms.Padding(3, 3, 3, 0);
            this.myShapeFlowLayout.MaximumSize = new System.Drawing.Size(500, 500);
            this.myShapeFlowLayout.MinimumSize = new System.Drawing.Size(120, 50);
            this.myShapeFlowLayout.Name = "myShapeFlowLayout";
            this.myShapeFlowLayout.Size = new System.Drawing.Size(377, 58);
            this.myShapeFlowLayout.TabIndex = 1;
            // 
            // contextMenuStrip
            // 
            this.contextMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.removeShapeToolStripMenuItem,
            this.editNameToolStripMenuItem});
            this.contextMenuStrip.Name = "contextMenuStrip";
            this.contextMenuStrip.Size = new System.Drawing.Size(164, 48);
            this.contextMenuStrip.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.ContextMenuStripItemClicked);
            // 
            // removeShapeToolStripMenuItem
            // 
            this.removeShapeToolStripMenuItem.Name = "removeShapeToolStripMenuItem";
            this.removeShapeToolStripMenuItem.Size = new System.Drawing.Size(163, 22);
            this.removeShapeToolStripMenuItem.Text = "Remove Shape";
            // 
            // editNameToolStripMenuItem
            // 
            this.editNameToolStripMenuItem.Name = "editNameToolStripMenuItem";
            this.editNameToolStripMenuItem.Size = new System.Drawing.Size(163, 22);
            this.editNameToolStripMenuItem.Text = "Edit Name";
            // 
            // CustomShapePane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.motherTableLayoutPanel);
            this.Name = "CustomShapePane";
            this.Size = new System.Drawing.Size(390, 459);
            this.Click += new System.EventHandler(this.CustomShapePaneClick);
            this.motherTableLayoutPanel.ResumeLayout(false);
            this.motherTableLayoutPanel.PerformLayout();
            this.myShapePanel.ResumeLayout(false);
            this.contextMenuStrip.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel motherTableLayoutPanel;
        private ExtendedPanel myShapePanel;
        private System.Windows.Forms.FlowLayoutPanel myShapeFlowLayout;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip;
        private System.Windows.Forms.ToolStripMenuItem removeShapeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem editNameToolStripMenuItem;

    }
}
