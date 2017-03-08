using System.Drawing;
using System.Windows.Forms;
using PowerPointLabs.Utils;

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
            this.moveShapeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.copyToToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.removeShapeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.flowlayoutContextMenuStrip = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.addCategoryToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.removeCategoryToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.renameCategoryToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.setAsDefaultToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importCategoryToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importShapesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.settingsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label1 = new System.Windows.Forms.Label();
            this.categoryBox = new System.Windows.Forms.ComboBox();
            this.flowPanelHolder = new System.Windows.Forms.Panel();
            this.myShapeFlowLayout = new PowerPointLabs.BufferedFlowLayoutPanel();
            this.singleShapeDownloadLink = new System.Windows.Forms.LinkLabel();
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
            this.moveShapeToolStripMenuItem,
            this.copyToToolStripMenuItem,
            this.removeShapeToolStripMenuItem});
            this.shapeContextMenuStrip.Name = "contextMenuStrip";
            this.shapeContextMenuStrip.Size = new System.Drawing.Size(68, 114);
            this.shapeContextMenuStrip.Opening += new System.ComponentModel.CancelEventHandler(this.ThumbnailContextMenuStripOpening);
            this.shapeContextMenuStrip.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.ThumbnailContextMenuStripItemClicked);
            // 
            // addToSlideToolStripMenuItem
            // 
            this.addToSlideToolStripMenuItem.Name = "addToSlideToolStripMenuItem";
            this.addToSlideToolStripMenuItem.Size = new System.Drawing.Size(67, 22);
            // 
            // editNameToolStripMenuItem
            // 
            this.editNameToolStripMenuItem.Name = "editNameToolStripMenuItem";
            this.editNameToolStripMenuItem.Size = new System.Drawing.Size(67, 22);
            // 
            // moveShapeToolStripMenuItem
            // 
            this.moveShapeToolStripMenuItem.Name = "moveShapeToolStripMenuItem";
            this.moveShapeToolStripMenuItem.Size = new System.Drawing.Size(67, 22);
            this.moveShapeToolStripMenuItem.Click += new System.EventHandler(this.MoveContextMenuStripOnEvent);
            this.moveShapeToolStripMenuItem.MouseEnter += new System.EventHandler(this.MoveContextMenuStripOnEvent);
            // 
            // copyToToolStripMenuItem
            // 
            this.copyToToolStripMenuItem.Name = "copyToToolStripMenuItem";
            this.copyToToolStripMenuItem.Size = new System.Drawing.Size(67, 22);
            this.copyToToolStripMenuItem.Click += new System.EventHandler(this.CopyContextMenuStripOnEvent);
            this.copyToToolStripMenuItem.MouseEnter += new System.EventHandler(this.CopyContextMenuStripOnEvent);
            // 
            // removeShapeToolStripMenuItem
            // 
            this.removeShapeToolStripMenuItem.Name = "removeShapeToolStripMenuItem";
            this.removeShapeToolStripMenuItem.Size = new System.Drawing.Size(67, 22);
            // 
            // flowlayoutContextMenuStrip
            // 
            this.flowlayoutContextMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.addCategoryToolStripMenuItem,
            this.removeCategoryToolStripMenuItem,
            this.renameCategoryToolStripMenuItem,
            this.setAsDefaultToolStripMenuItem,
            this.importCategoryToolStripMenuItem,
            this.importShapesToolStripMenuItem,
            this.settingsToolStripMenuItem});
            this.flowlayoutContextMenuStrip.Name = "flowlayoutContextMenuStrip";
            this.flowlayoutContextMenuStrip.Size = new System.Drawing.Size(68, 158);
            this.flowlayoutContextMenuStrip.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.FlowlayoutContextMenuStripItemClicked);
            // 
            // addCategoryToolStripMenuItem
            // 
            this.addCategoryToolStripMenuItem.Name = "addCategoryToolStripMenuItem";
            this.addCategoryToolStripMenuItem.Size = new System.Drawing.Size(67, 22);
            // 
            // removeCategoryToolStripMenuItem
            // 
            this.removeCategoryToolStripMenuItem.Name = "removeCategoryToolStripMenuItem";
            this.removeCategoryToolStripMenuItem.Size = new System.Drawing.Size(67, 22);
            // 
            // renameCategoryToolStripMenuItem
            // 
            this.renameCategoryToolStripMenuItem.Name = "renameCategoryToolStripMenuItem";
            this.renameCategoryToolStripMenuItem.Size = new System.Drawing.Size(67, 22);
            // 
            // setAsDefaultToolStripMenuItem
            // 
            this.setAsDefaultToolStripMenuItem.Name = "setAsDefaultToolStripMenuItem";
            this.setAsDefaultToolStripMenuItem.Size = new System.Drawing.Size(67, 22);
            // 
            // importCategoryToolStripMenuItem
            // 
            this.importCategoryToolStripMenuItem.Name = "importCategoryToolStripMenuItem";
            this.importCategoryToolStripMenuItem.Size = new System.Drawing.Size(67, 22);
            // 
            // importShapesToolStripMenuItem
            // 
            this.importShapesToolStripMenuItem.Name = "importShapesToolStripMenuItem";
            this.importShapesToolStripMenuItem.Size = new System.Drawing.Size(67, 22);
            // 
            // settingsToolStripMenuItem
            // 
            this.settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
            this.settingsToolStripMenuItem.Size = new System.Drawing.Size(67, 22);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(19, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(60, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Category";
            // 
            // categoryBox
            // 
            this.categoryBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.categoryBox.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.categoryBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.categoryBox.Location = new System.Drawing.Point(88, 18);
            this.categoryBox.Name = "categoryBox";
            this.categoryBox.Size = new System.Drawing.Size(314, 21);
            this.categoryBox.TabIndex = 3;
            this.categoryBox.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.CategoryBoxOwnerDraw);
            this.categoryBox.SelectedIndexChanged += new System.EventHandler(this.CategoryBoxSelectedIndexChanged);
            // 
            // flowPanelHolder
            // 
            this.flowPanelHolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.flowPanelHolder.Controls.Add(this.myShapeFlowLayout);
            this.flowPanelHolder.Location = new System.Drawing.Point(3, 47);
            this.flowPanelHolder.Name = "flowPanelHolder";
            this.flowPanelHolder.Size = new System.Drawing.Size(415, 516);
            this.flowPanelHolder.TabIndex = 4;
            // 
            // myShapeFlowLayout
            // 
            this.myShapeFlowLayout.AutoScroll = true;
            this.myShapeFlowLayout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.myShapeFlowLayout.ContextMenuStrip = this.flowlayoutContextMenuStrip;
            this.myShapeFlowLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.myShapeFlowLayout.Location = new System.Drawing.Point(0, 0);
            this.myShapeFlowLayout.MinimumSize = new System.Drawing.Size(120, 54);
            this.myShapeFlowLayout.Name = "myShapeFlowLayout";
            this.myShapeFlowLayout.Size = new System.Drawing.Size(415, 516);
            this.myShapeFlowLayout.TabIndex = 2;
            // 
            // singleShapeDownloadLink
            // 
            this.singleShapeDownloadLink.AutoSize = true;
            this.singleShapeDownloadLink.Location = new System.Drawing.Point(3, 566);
            this.singleShapeDownloadLink.Name = "singleShapeDownloadLink";
            this.singleShapeDownloadLink.Size = new System.Drawing.Size(114, 13);
            this.singleShapeDownloadLink.TabIndex = 5;
            this.singleShapeDownloadLink.TabStop = true;
            this.singleShapeDownloadLink.Text = "Find more shapes here";
            this.singleShapeDownloadLink.Visible = false;
            this.singleShapeDownloadLink.VisitedLinkColor = System.Drawing.Color.Blue;
            // 
            // CustomShapePane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ContextMenuStrip = this.flowlayoutContextMenuStrip;
            this.Controls.Add(this.singleShapeDownloadLink);
            this.Controls.Add(this.flowPanelHolder);
            this.Controls.Add(this.categoryBox);
            this.Controls.Add(this.label1);
            this.Name = "CustomShapePane";
            this.Size = new System.Drawing.Size(421, 598);
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
            Location = new Point((int)(81 * Utils.Graphics.GetDpiScale()), 
                                 (int)(11 * Utils.Graphics.GetDpiScale())),
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
            Location = new Point((int)(11 * Utils.Graphics.GetDpiScale()),
                                 (int)(41 * Utils.Graphics.GetDpiScale())),
            Name = "noShapeLabel",
            Size = new Size(242, 24),
            Text = TextCollection.CustomShapeNoShapeTextSecondLine
        };

        private readonly Panel _noShapePanel = new Panel
        {
            AutoSize = true,
            Name = "noShapePanel",
            Size = new Size(392, 100),
            Margin = new Padding(0, 0, 0, 0)
        };
        private ToolStripMenuItem addToSlideToolStripMenuItem;
        private ContextMenuStrip flowlayoutContextMenuStrip;
        private ToolStripMenuItem settingsToolStripMenuItem;
        private Label label1;
        private ComboBox categoryBox;
        private Panel flowPanelHolder;
        private ToolStripMenuItem addCategoryToolStripMenuItem;
        private ToolStripMenuItem moveShapeToolStripMenuItem;
        private ToolStripMenuItem removeCategoryToolStripMenuItem;
        private ToolStripMenuItem renameCategoryToolStripMenuItem;
        private ToolStripMenuItem copyToToolStripMenuItem;
        private ToolStripMenuItem setAsDefaultToolStripMenuItem;
        private ToolStripMenuItem importCategoryToolStripMenuItem;
        private BufferedFlowLayoutPanel myShapeFlowLayout;
        private ToolStripMenuItem importShapesToolStripMenuItem;
        private LinkLabel singleShapeDownloadLink;
    }
}
