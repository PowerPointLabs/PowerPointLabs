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
            this.shapeContextMenuStrip = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.addToSlideToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.editNameToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.removeShapeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.flowlayoutContextMenuStrip = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.settingsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.myShapeFlowLayout = new System.Windows.Forms.FlowLayoutPanel();
            this.label1 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.flowPanelHolder = new System.Windows.Forms.Panel();
            this.shapeContextMenuStrip.SuspendLayout();
            this.flowlayoutContextMenuStrip.SuspendLayout();
            this.flowPanelHolder.SuspendLayout();
            this.SuspendLayout();
            // 
            // shapeContextMenuStrip
            // 
            this.shapeContextMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.addToSlideToolStripMenuItem,
            this.editNameToolStripMenuItem,
            this.removeShapeToolStripMenuItem});
            this.shapeContextMenuStrip.Name = "contextMenuStrip";
            this.shapeContextMenuStrip.Size = new System.Drawing.Size(164, 70);
            this.shapeContextMenuStrip.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.ThumbnailContextMenuStripItemClicked);
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
            // flowlayoutContextMenuStrip
            // 
            this.flowlayoutContextMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.settingsToolStripMenuItem});
            this.flowlayoutContextMenuStrip.Name = "flowlayoutContextMenuStrip";
            this.flowlayoutContextMenuStrip.Size = new System.Drawing.Size(123, 26);
            this.flowlayoutContextMenuStrip.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.FlowlayoutContextMenuStripItemClicked);
            // 
            // settingsToolStripMenuItem
            // 
            this.settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
            this.settingsToolStripMenuItem.Size = new System.Drawing.Size(122, 22);
            this.settingsToolStripMenuItem.Text = "Settings";
            // 
            // myShapeFlowLayout
            // 
            this.myShapeFlowLayout.AutoScroll = true;
            this.myShapeFlowLayout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.myShapeFlowLayout.ContextMenuStrip = this.flowlayoutContextMenuStrip;
            this.myShapeFlowLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.myShapeFlowLayout.Location = new System.Drawing.Point(0, 0);
            this.myShapeFlowLayout.Margin = new System.Windows.Forms.Padding(0);
            this.myShapeFlowLayout.MinimumSize = new System.Drawing.Size(120, 50);
            this.myShapeFlowLayout.Name = "myShapeFlowLayout";
            this.myShapeFlowLayout.Size = new System.Drawing.Size(411, 476);
            this.myShapeFlowLayout.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "Category";
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(72, 17);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(316, 20);
            this.comboBox1.TabIndex = 3;
            // 
            // flowPanelHolder
            // 
            this.flowPanelHolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.flowPanelHolder.Controls.Add(this.myShapeFlowLayout);
            this.flowPanelHolder.Location = new System.Drawing.Point(3, 43);
            this.flowPanelHolder.Name = "flowPanelHolder";
            this.flowPanelHolder.Size = new System.Drawing.Size(411, 476);
            this.flowPanelHolder.TabIndex = 4;
            // 
            // CustomShapePane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ContextMenuStrip = this.flowlayoutContextMenuStrip;
            this.Controls.Add(this.flowPanelHolder);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.label1);
            this.Name = "CustomShapePane";
            this.Size = new System.Drawing.Size(417, 552);
            this.Click += new System.EventHandler(this.CustomShapePaneClick);
            this.shapeContextMenuStrip.ResumeLayout(false);
            this.flowlayoutContextMenuStrip.ResumeLayout(false);
            this.flowPanelHolder.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ContextMenuStrip shapeContextMenuStrip;
        private System.Windows.Forms.ToolStripMenuItem removeShapeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem editNameToolStripMenuItem;

        private readonly Label _noShapeLabelFirstLine = new Label
        {
            AutoSize = true,
            Font =
                new Font("Arial", 15.75F, FontStyle.Bold, GraphicsUnit.Point,
                         0),
            ForeColor = SystemColors.ButtonShadow,
            Location = new Point(81, 11),
            Name = "noShapeLabel",
            Size = new Size(212, 24),
            Text = TextCollection.CustomShapeNoShapeTextFirstLine
        };

        private readonly Label _noShapeLabelSecondLine = new Label
        {
            AutoSize = true,
            Font =
                new Font("Arial", 10F, FontStyle.Bold, GraphicsUnit.Point,
                         0),
            ForeColor = SystemColors.ButtonShadow,
            Location = new Point(11, 41),
            Name = "noShapeLabel",
            Size = new Size(242, 24),
            Text = TextCollection.CustomShapeNoShapeTextSecondLine
        };

        private readonly Panel _noShapePanel = new Panel
        {
            Name = "noShapePanel",
            Size = new Size(392, 100),
            Margin = new Padding(0, 0, 0, 0)
        };
        private ToolStripMenuItem addToSlideToolStripMenuItem;
        private ContextMenuStrip flowlayoutContextMenuStrip;
        private ToolStripMenuItem settingsToolStripMenuItem;
        private FlowLayoutPanel myShapeFlowLayout;
        private Label label1;
        private ComboBox comboBox1;
        private Panel flowPanelHolder;
    }
}
