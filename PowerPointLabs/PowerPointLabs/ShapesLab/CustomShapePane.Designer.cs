using System.Drawing;
using System.Windows.Forms;

using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ShapesLab
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
            this.addShapeButton = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.shapeContextMenuStrip.SuspendLayout();
            this.flowlayoutContextMenuStrip.SuspendLayout();
            this.flowPanelHolder.SuspendLayout();
            this.SuspendLayout();
            // 
            // shapeContextMenuStrip
            // 
            this.shapeContextMenuStrip.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.shapeContextMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.addToSlideToolStripMenuItem,
            this.editNameToolStripMenuItem,
            this.moveShapeToolStripMenuItem,
            this.copyToToolStripMenuItem,
            this.removeShapeToolStripMenuItem});
            this.shapeContextMenuStrip.Name = "contextMenuStrip";
            this.shapeContextMenuStrip.Size = new System.Drawing.Size(76, 114);
            this.shapeContextMenuStrip.Opening += new System.ComponentModel.CancelEventHandler(this.ThumbnailContextMenuStripOpening);
            this.shapeContextMenuStrip.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.ThumbnailContextMenuStripItemClicked);
            // 
            // addToSlideToolStripMenuItem
            // 
            this.addToSlideToolStripMenuItem.Name = "addToSlideToolStripMenuItem";
            this.addToSlideToolStripMenuItem.Size = new System.Drawing.Size(75, 22);
            // 
            // editNameToolStripMenuItem
            // 
            this.editNameToolStripMenuItem.Name = "editNameToolStripMenuItem";
            this.editNameToolStripMenuItem.Size = new System.Drawing.Size(75, 22);
            // 
            // moveShapeToolStripMenuItem
            // 
            this.moveShapeToolStripMenuItem.Name = "moveShapeToolStripMenuItem";
            this.moveShapeToolStripMenuItem.Size = new System.Drawing.Size(75, 22);
            this.moveShapeToolStripMenuItem.Click += new System.EventHandler(this.MoveContextMenuStripOnEvent);
            this.moveShapeToolStripMenuItem.MouseEnter += new System.EventHandler(this.MoveContextMenuStripOnEvent);
            // 
            // copyToToolStripMenuItem
            // 
            this.copyToToolStripMenuItem.Name = "copyToToolStripMenuItem";
            this.copyToToolStripMenuItem.Size = new System.Drawing.Size(75, 22);
            this.copyToToolStripMenuItem.Click += new System.EventHandler(this.CopyContextMenuStripOnEvent);
            this.copyToToolStripMenuItem.MouseEnter += new System.EventHandler(this.CopyContextMenuStripOnEvent);
            // 
            // removeShapeToolStripMenuItem
            // 
            this.removeShapeToolStripMenuItem.Name = "removeShapeToolStripMenuItem";
            this.removeShapeToolStripMenuItem.Size = new System.Drawing.Size(75, 22);
            // 
            // flowlayoutContextMenuStrip
            // 
            this.flowlayoutContextMenuStrip.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.flowlayoutContextMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.addCategoryToolStripMenuItem,
            this.removeCategoryToolStripMenuItem,
            this.renameCategoryToolStripMenuItem,
            this.setAsDefaultToolStripMenuItem,
            this.importCategoryToolStripMenuItem,
            this.importShapesToolStripMenuItem,
            this.settingsToolStripMenuItem});
            this.flowlayoutContextMenuStrip.Name = "flowlayoutContextMenuStrip";
            this.flowlayoutContextMenuStrip.Size = new System.Drawing.Size(76, 158);
            this.flowlayoutContextMenuStrip.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.FlowlayoutContextMenuStripItemClicked);
            // 
            // addCategoryToolStripMenuItem
            // 
            this.addCategoryToolStripMenuItem.Name = "addCategoryToolStripMenuItem";
            this.addCategoryToolStripMenuItem.Size = new System.Drawing.Size(75, 22);
            // 
            // removeCategoryToolStripMenuItem
            // 
            this.removeCategoryToolStripMenuItem.Name = "removeCategoryToolStripMenuItem";
            this.removeCategoryToolStripMenuItem.Size = new System.Drawing.Size(75, 22);
            // 
            // renameCategoryToolStripMenuItem
            // 
            this.renameCategoryToolStripMenuItem.Name = "renameCategoryToolStripMenuItem";
            this.renameCategoryToolStripMenuItem.Size = new System.Drawing.Size(75, 22);
            // 
            // setAsDefaultToolStripMenuItem
            // 
            this.setAsDefaultToolStripMenuItem.Name = "setAsDefaultToolStripMenuItem";
            this.setAsDefaultToolStripMenuItem.Size = new System.Drawing.Size(75, 22);
            // 
            // importCategoryToolStripMenuItem
            // 
            this.importCategoryToolStripMenuItem.Name = "importCategoryToolStripMenuItem";
            this.importCategoryToolStripMenuItem.Size = new System.Drawing.Size(75, 22);
            // 
            // importShapesToolStripMenuItem
            // 
            this.importShapesToolStripMenuItem.Name = "importShapesToolStripMenuItem";
            this.importShapesToolStripMenuItem.Size = new System.Drawing.Size(75, 22);
            // 
            // settingsToolStripMenuItem
            // 
            this.settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
            this.settingsToolStripMenuItem.Size = new System.Drawing.Size(75, 22);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(38, 162);
            this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(111, 26);
            this.label1.TabIndex = 2;
            this.label1.Text = "Category";
            // 
            // categoryBox
            // 
            this.categoryBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.categoryBox.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.categoryBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.categoryBox.Location = new System.Drawing.Point(176, 158);
            this.categoryBox.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.categoryBox.Name = "categoryBox";
            this.categoryBox.Size = new System.Drawing.Size(624, 32);
            this.categoryBox.TabIndex = 3;
            this.categoryBox.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.CategoryBoxOwnerDraw);
            this.categoryBox.SelectedIndexChanged += new System.EventHandler(this.CategoryBoxSelectedIndexChanged);
            // 
            // flowPanelHolder
            // 
            this.flowPanelHolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.flowPanelHolder.Controls.Add(this.myShapeFlowLayout);
            this.flowPanelHolder.Location = new System.Drawing.Point(6, 212);
            this.flowPanelHolder.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.flowPanelHolder.Name = "flowPanelHolder";
            this.flowPanelHolder.Size = new System.Drawing.Size(830, 871);
            this.flowPanelHolder.TabIndex = 4;
            // 
            // myShapeFlowLayout
            // 
            this.myShapeFlowLayout.AutoScroll = true;
            this.myShapeFlowLayout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.myShapeFlowLayout.ContextMenuStrip = this.flowlayoutContextMenuStrip;
            this.myShapeFlowLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.myShapeFlowLayout.Location = new System.Drawing.Point(0, 0);
            this.myShapeFlowLayout.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.myShapeFlowLayout.MinimumSize = new System.Drawing.Size(240, 104);
            this.myShapeFlowLayout.Name = "myShapeFlowLayout";
            this.myShapeFlowLayout.Size = new System.Drawing.Size(830, 871);
            this.myShapeFlowLayout.TabIndex = 2;
            // 
            // singleShapeDownloadLink
            // 
            this.singleShapeDownloadLink.AutoSize = true;
            this.singleShapeDownloadLink.Location = new System.Drawing.Point(6, 1088);
            this.singleShapeDownloadLink.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.singleShapeDownloadLink.Name = "singleShapeDownloadLink";
            this.singleShapeDownloadLink.Size = new System.Drawing.Size(233, 25);
            this.singleShapeDownloadLink.TabIndex = 5;
            this.singleShapeDownloadLink.TabStop = true;
            this.singleShapeDownloadLink.Text = "Find more shapes here";
            this.singleShapeDownloadLink.Visible = false;
            this.singleShapeDownloadLink.VisitedLinkColor = System.Drawing.Color.Blue;
            // 
            // addShapeButton
            // 
            this.addShapeButton.BackColor = System.Drawing.SystemColors.Control;
            this.addShapeButton.BackgroundImage = global::PowerPointLabs.Properties.Resources.AddToCustomShapes;
            this.addShapeButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.addShapeButton.FlatAppearance.BorderColor = System.Drawing.Color.Black;
            this.addShapeButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.addShapeButton.Location = new System.Drawing.Point(8, 13);
            this.addShapeButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.addShapeButton.Name = "addShapeButton";
            this.addShapeButton.Size = new System.Drawing.Size(88, 88);
            this.addShapeButton.TabIndex = 6;
            this.addShapeButton.UseVisualStyleBackColor = false;
            this.addShapeButton.Click += new System.EventHandler(this.AddShapeButton_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(26, 110);
            this.label2.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(55, 26);
            this.label2.TabIndex = 7;
            this.label2.Text = "Add";
            // 
            // CustomShapePane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ContextMenuStrip = this.flowlayoutContextMenuStrip;
            this.Controls.Add(this.label2);
            this.Controls.Add(this.addShapeButton);
            this.Controls.Add(this.singleShapeDownloadLink);
            this.Controls.Add(this.flowPanelHolder);
            this.Controls.Add(this.categoryBox);
            this.Controls.Add(this.label1);
            this.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.Name = "CustomShapePane";
            this.Size = new System.Drawing.Size(842, 1150);
            this.Click += new System.EventHandler(this.CustomShapePaneClick);
            this.MouseMove += new System.Windows.Forms.MouseEventHandler(this.CustomShapePane_MouseMove);
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
            Location = new Point((int)(81 * Utils.GraphicsUtil.GetDpiScale()),
                                 (int)(11 * Utils.GraphicsUtil.GetDpiScale())),
            Name = "noShapeLabel",
            Size = new Size(212, 24),
            Text = ShapesLabText.ErrorNoShapeTextFirstLine
        };

        private readonly Label _noShapeLabelSecondLine = new Label
        {
            AutoSize = true,
            Font =
                new Font("Arial", 10F, FontStyle.Bold, GraphicsUnit.Point,
                         0),
            ForeColor = SystemColors.ButtonShadow,
            Location = new Point((int)(11 * Utils.GraphicsUtil.GetDpiScale()),
                                 (int)(41 * Utils.GraphicsUtil.GetDpiScale())),
            Name = "noShapeLabel",
            Size = new Size(242, 24),
            Text = ShapesLabText.ErrorNoShapeTextSecondLine
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
        private Button addShapeButton;
        private Label label2;
        private ToolTip toolTip1;
    }
}
