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
            this.tabControl = new System.Windows.Forms.TabControl();
            this.defaultPage = new System.Windows.Forms.TabPage();
            this.myShapeFlowLayout = new System.Windows.Forms.FlowLayoutPanel();
            this.flowlayoutContextMenuStrip = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.settingsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.addMorePageButton = new System.Windows.Forms.TabPage();
            this.shapeContextMenuStrip.SuspendLayout();
            this.tabControl.SuspendLayout();
            this.defaultPage.SuspendLayout();
            this.flowlayoutContextMenuStrip.SuspendLayout();
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
            // tabControl
            // 
            this.tabControl.Controls.Add(this.defaultPage);
            this.tabControl.Controls.Add(this.addMorePageButton);
            this.tabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl.Location = new System.Drawing.Point(0, 0);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(417, 499);
            this.tabControl.TabIndex = 5;
            // 
            // defaultPage
            // 
            this.defaultPage.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.defaultPage.Controls.Add(this.myShapeFlowLayout);
            this.defaultPage.Location = new System.Drawing.Point(4, 22);
            this.defaultPage.Name = "defaultPage";
            this.defaultPage.Padding = new System.Windows.Forms.Padding(3);
            this.defaultPage.Size = new System.Drawing.Size(409, 473);
            this.defaultPage.TabIndex = 0;
            this.defaultPage.Text = "My Saved Shapes";
            this.defaultPage.UseVisualStyleBackColor = true;
            // 
            // myShapeFlowLayout
            // 
            this.myShapeFlowLayout.AutoScroll = true;
            this.myShapeFlowLayout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.myShapeFlowLayout.ContextMenuStrip = this.flowlayoutContextMenuStrip;
            this.myShapeFlowLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.myShapeFlowLayout.Location = new System.Drawing.Point(3, 3);
            this.myShapeFlowLayout.Margin = new System.Windows.Forms.Padding(0);
            this.myShapeFlowLayout.MaximumSize = new System.Drawing.Size(700, 500);
            this.myShapeFlowLayout.MinimumSize = new System.Drawing.Size(120, 50);
            this.myShapeFlowLayout.Name = "myShapeFlowLayout";
            this.myShapeFlowLayout.Size = new System.Drawing.Size(399, 463);
            this.myShapeFlowLayout.TabIndex = 1;
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
            // addMorePageButton
            // 
            this.addMorePageButton.Location = new System.Drawing.Point(4, 22);
            this.addMorePageButton.Name = "addMorePageButton";
            this.addMorePageButton.Padding = new System.Windows.Forms.Padding(3);
            this.addMorePageButton.Size = new System.Drawing.Size(409, 473);
            this.addMorePageButton.TabIndex = 1;
            this.addMorePageButton.Text = "Add...";
            this.addMorePageButton.UseVisualStyleBackColor = true;
            this.addMorePageButton.Click += new System.EventHandler(this.addMorePageButton_Click);
            // 
            // CustomShapePane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ContextMenuStrip = this.flowlayoutContextMenuStrip;
            this.Controls.Add(this.tabControl);
            this.Name = "CustomShapePane";
            this.Size = new System.Drawing.Size(417, 499);
            this.Click += new System.EventHandler(this.CustomShapePaneClick);
            this.shapeContextMenuStrip.ResumeLayout(false);
            this.tabControl.ResumeLayout(false);
            this.defaultPage.ResumeLayout(false);
            this.flowlayoutContextMenuStrip.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ContextMenuStrip shapeContextMenuStrip;
        private System.Windows.Forms.ToolStripMenuItem removeShapeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem editNameToolStripMenuItem;
        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage defaultPage;
        private System.Windows.Forms.FlowLayoutPanel myShapeFlowLayout;

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
        private TabPage addMorePageButton;
    }
}
