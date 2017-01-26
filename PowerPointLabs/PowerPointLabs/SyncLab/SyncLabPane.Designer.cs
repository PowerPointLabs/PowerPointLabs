using System.Drawing;
using System.Windows.Forms;

namespace PowerPointLabs
{
    partial class SyncLabPane
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
            this.copyButton = new System.Windows.Forms.Button();
            this.pasteButton = new System.Windows.Forms.Button();
            this.copyLabel = new System.Windows.Forms.Label();
            this.pasteLabel = new System.Windows.Forms.Label();
            this.syncLabListBox = new PowerPointLabs.SyncLab.SyncLabListBox();
            this.SuspendLayout();
            // 
            // copyButton
            // 
            this.copyButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.copyButton.Image = global::PowerPointLabs.Properties.Resources.LineColor_icon;
            this.copyButton.Location = new System.Drawing.Point(12, 12);
            this.copyButton.Name = "copyButton";
            this.copyButton.Size = new System.Drawing.Size(45, 45);
            this.copyButton.TabIndex = 27;
            this.copyButton.UseVisualStyleBackColor = true;
            this.copyButton.Click += new System.EventHandler(this.CopyButton_Click);
            // 
            // pasteButton
            // 
            this.pasteButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pasteButton.Image = global::PowerPointLabs.Properties.Resources.FillColor_icon;
            this.pasteButton.Location = new System.Drawing.Point(69, 12);
            this.pasteButton.Name = "pasteButton";
            this.pasteButton.Size = new System.Drawing.Size(45, 45);
            this.pasteButton.TabIndex = 26;
            this.pasteButton.UseVisualStyleBackColor = true;
            this.pasteButton.Click += new System.EventHandler(this.PasteButton_Click);
            // 
            // copyLabel
            // 
            this.copyLabel.BackColor = System.Drawing.Color.Transparent;
            this.copyLabel.Location = new System.Drawing.Point(11, 55);
            this.copyLabel.Name = "copyLabel";
            this.copyLabel.Size = new System.Drawing.Size(48, 21);
            this.copyLabel.TabIndex = 28;
            this.copyLabel.Text = "Copy";
            this.copyLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // pasteLabel
            // 
            this.pasteLabel.BackColor = System.Drawing.Color.Transparent;
            this.pasteLabel.Location = new System.Drawing.Point(67, 55);
            this.pasteLabel.Name = "pasteLabel";
            this.pasteLabel.Size = new System.Drawing.Size(48, 21);
            this.pasteLabel.TabIndex = 29;
            this.pasteLabel.Text = "Paste";
            this.pasteLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // syncLabListBox
            // 
            this.syncLabListBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.syncLabListBox.Location = new System.Drawing.Point(12, 79);
            this.syncLabListBox.Name = "syncLabListBox";
            this.syncLabListBox.Size = new System.Drawing.Size(276, 801);
            this.syncLabListBox.TabIndex = 30;
            this.syncLabListBox.Text = "syncLabListBox";
            this.syncLabListBox.UseCompatibleStateImageBehavior = false;
            // 
            // SyncLabPane
            // 
            this.Controls.Add(this.syncLabListBox);
            this.Controls.Add(this.copyButton);
            this.Controls.Add(this.pasteButton);
            this.Controls.Add(this.copyLabel);
            this.Controls.Add(this.pasteLabel);
            this.Name = "SyncLabPane";
            this.Size = new System.Drawing.Size(300, 883);
            this.ResumeLayout(false);

        }

        #endregion

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
        private Button copyButton;
        private Button pasteButton;
        private Label copyLabel;
        private Label pasteLabel;
        private SyncLab.SyncLabListBox syncLabListBox;
    }
}
