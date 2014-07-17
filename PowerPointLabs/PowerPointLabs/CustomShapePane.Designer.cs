using System.Drawing;
using System.Windows.Forms;
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
            this.contextMenuStrip = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.addToSlideToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.editNameToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.removeShapeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.myShapeFlowLayout = new System.Windows.Forms.FlowLayoutPanel();
            this.contextMenuStrip.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.SuspendLayout();
            // 
            // contextMenuStrip
            // 
            this.contextMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.addToSlideToolStripMenuItem,
            this.editNameToolStripMenuItem,
            this.removeShapeToolStripMenuItem});
            this.contextMenuStrip.Name = "contextMenuStrip";
            this.contextMenuStrip.Size = new System.Drawing.Size(164, 70);
            this.contextMenuStrip.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.ContextMenuStripItemClicked);
            // 
            // addToSlideToolStripMenuItem
            // 
            this.addToSlideToolStripMenuItem.Name = "addToSlideToolStripMenuItem";
            this.addToSlideToolStripMenuItem.Size = new System.Drawing.Size(163, 22);
            this.addToSlideToolStripMenuItem.Text = "Add to Slide";
            // 
            // editNameToolStripMenuItem
            // 
            this.editNameToolStripMenuItem.Name = "editNameToolStripMenuItem";
            this.editNameToolStripMenuItem.Size = new System.Drawing.Size(163, 22);
            this.editNameToolStripMenuItem.Text = "Edit Name";
            // 
            // removeShapeToolStripMenuItem
            // 
            this.removeShapeToolStripMenuItem.Name = "removeShapeToolStripMenuItem";
            this.removeShapeToolStripMenuItem.Size = new System.Drawing.Size(163, 22);
            this.removeShapeToolStripMenuItem.Text = "Remove Shape";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(417, 499);
            this.tabControl1.TabIndex = 5;
            // 
            // tabPage1
            // 
            this.tabPage1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.tabPage1.Controls.Add(this.myShapeFlowLayout);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(409, 473);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "My Saved Shapes";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // myShapeFlowLayout
            // 
            this.myShapeFlowLayout.AutoScroll = true;
            this.myShapeFlowLayout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.myShapeFlowLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.myShapeFlowLayout.Location = new System.Drawing.Point(3, 3);
            this.myShapeFlowLayout.Margin = new System.Windows.Forms.Padding(0);
            this.myShapeFlowLayout.MaximumSize = new System.Drawing.Size(700, 500);
            this.myShapeFlowLayout.MinimumSize = new System.Drawing.Size(120, 50);
            this.myShapeFlowLayout.Name = "myShapeFlowLayout";
            this.myShapeFlowLayout.Size = new System.Drawing.Size(399, 463);
            this.myShapeFlowLayout.TabIndex = 1;
            // 
            // CustomShapePane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl1);
            this.Name = "CustomShapePane";
            this.Size = new System.Drawing.Size(417, 499);
            this.Click += new System.EventHandler(this.CustomShapePaneClick);
            this.contextMenuStrip.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ContextMenuStrip contextMenuStrip;
        private System.Windows.Forms.ToolStripMenuItem removeShapeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem editNameToolStripMenuItem;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.FlowLayoutPanel myShapeFlowLayout;

        private readonly Label _noShapeLabel = new Label
        {
            AutoSize = true,
            Font =
                new Font("Arial", 15.75F, FontStyle.Bold, GraphicsUnit.Point,
                         0),
            ForeColor = SystemColors.ButtonShadow,
            Location = new Point(81, 11),
            Name = "noShapeLabel",
            Size = new Size(212, 24),
            Text = TextCollection.CustomShapeNoShapeText
        };

        private readonly Panel _noShapePanel = new Panel
        {
            Name = "noShapePanel",
            Size = new Size(362, 50),
            Margin = new Padding(0, 0, 0, 0)
        };
        private ToolStripMenuItem addToSlideToolStripMenuItem;
    }
}
