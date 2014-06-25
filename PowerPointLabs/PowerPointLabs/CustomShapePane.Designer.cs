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
            this.motherTableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.mediaIconPanel = new Stepi.UI.ExtendedPanel();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel6 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.myShapePanel = new Stepi.UI.ExtendedPanel();
            this.searchBox = new System.Windows.Forms.TextBox();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.panel5 = new System.Windows.Forms.Panel();
            this.panel7 = new System.Windows.Forms.Panel();
            this.myShapeFlowLayout = new System.Windows.Forms.FlowLayoutPanel();
            this.motherTableLayoutPanel.SuspendLayout();
            this.mediaIconPanel.SuspendLayout();
            this.flowLayoutPanel1.SuspendLayout();
            this.myShapePanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // motherTableLayoutPanel
            // 
            this.motherTableLayoutPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.motherTableLayoutPanel.AutoScroll = true;
            this.motherTableLayoutPanel.ColumnCount = 1;
            this.motherTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.motherTableLayoutPanel.Controls.Add(this.mediaIconPanel, 0, 1);
            this.motherTableLayoutPanel.Controls.Add(this.myShapePanel, 0, 0);
            this.motherTableLayoutPanel.Location = new System.Drawing.Point(3, 25);
            this.motherTableLayoutPanel.Margin = new System.Windows.Forms.Padding(0);
            this.motherTableLayoutPanel.MaximumSize = new System.Drawing.Size(500, 336);
            this.motherTableLayoutPanel.Name = "motherTableLayoutPanel";
            this.motherTableLayoutPanel.RowCount = 3;
            this.motherTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.motherTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.motherTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 1F));
            this.motherTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.motherTableLayoutPanel.Size = new System.Drawing.Size(250, 336);
            this.motherTableLayoutPanel.TabIndex = 0;
            // 
            // mediaIconPanel
            // 
            this.mediaIconPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.mediaIconPanel.BorderColor = System.Drawing.Color.Gray;
            this.mediaIconPanel.CaptionColorOne = System.Drawing.Color.White;
            this.mediaIconPanel.CaptionColorTwo = System.Drawing.Color.FromArgb(((int)(((byte)(155)))), ((int)(((byte)(255)))), ((int)(((byte)(165)))), ((int)(((byte)(0)))));
            this.mediaIconPanel.CaptionFont = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.mediaIconPanel.CaptionSize = 21;
            this.mediaIconPanel.CaptionText = "Media Icons";
            this.mediaIconPanel.CaptionTextColor = System.Drawing.Color.Black;
            this.mediaIconPanel.Controls.Add(this.flowLayoutPanel1);
            this.mediaIconPanel.DirectionCtrlColor = System.Drawing.Color.DarkGray;
            this.mediaIconPanel.DirectionCtrlHoverColor = System.Drawing.Color.Orange;
            this.mediaIconPanel.Location = new System.Drawing.Point(0, 213);
            this.mediaIconPanel.Margin = new System.Windows.Forms.Padding(0);
            this.mediaIconPanel.Name = "mediaIconPanel";
            this.mediaIconPanel.Size = new System.Drawing.Size(250, 112);
            this.mediaIconPanel.State = Stepi.UI.ExtendedPanelState.Collapsed;
            this.mediaIconPanel.TabIndex = 1;
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.flowLayoutPanel1.Controls.Add(this.panel1);
            this.flowLayoutPanel1.Controls.Add(this.panel2);
            this.flowLayoutPanel1.Controls.Add(this.panel6);
            this.flowLayoutPanel1.Controls.Add(this.panel3);
            this.flowLayoutPanel1.Location = new System.Drawing.Point(3, 23);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(244, 86);
            this.flowLayoutPanel1.TabIndex = 1;
            // 
            // panel1
            // 
            this.panel1.BackgroundImage = global::PowerPointLabs.Properties.Resources.Play;
            this.panel1.Location = new System.Drawing.Point(3, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(32, 32);
            this.panel1.TabIndex = 0;
            this.panel1.DoubleClick += new System.EventHandler(this.PanelDoubleClick);
            // 
            // panel2
            // 
            this.panel2.BackgroundImage = global::PowerPointLabs.Properties.Resources.Pause;
            this.panel2.Location = new System.Drawing.Point(41, 3);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(32, 32);
            this.panel2.TabIndex = 1;
            // 
            // panel6
            // 
            this.panel6.BackgroundImage = global::PowerPointLabs.Properties.Resources.Record;
            this.panel6.Location = new System.Drawing.Point(79, 3);
            this.panel6.Name = "panel6";
            this.panel6.Size = new System.Drawing.Size(32, 32);
            this.panel6.TabIndex = 2;
            // 
            // panel3
            // 
            this.panel3.BackgroundImage = global::PowerPointLabs.Properties.Resources.Stop;
            this.panel3.Location = new System.Drawing.Point(117, 3);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(32, 32);
            this.panel3.TabIndex = 2;
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
            // searchBox
            // 
            this.searchBox.Dock = System.Windows.Forms.DockStyle.Top;
            this.searchBox.Font = new System.Drawing.Font("Arial Narrow", 9F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.searchBox.ForeColor = System.Drawing.Color.Silver;
            this.searchBox.Location = new System.Drawing.Point(0, 0);
            this.searchBox.Name = "searchBox";
            this.searchBox.Size = new System.Drawing.Size(257, 21);
            this.searchBox.TabIndex = 3;
            this.searchBox.Text = "Search shapes...";
            this.searchBox.Enter += new System.EventHandler(this.SearchBoxEnter);
            this.searchBox.Leave += new System.EventHandler(this.SearchBoxLeave);
            this.searchBox.MouseUp += new System.Windows.Forms.MouseEventHandler(this.SearchBoxMouseUp);
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(1, 447);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(161, 12);
            this.linkLabel1.TabIndex = 6;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "Link to online shape store";
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
            // CustomShapePane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.motherTableLayoutPanel);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.searchBox);
            this.Name = "CustomShapePane";
            this.Size = new System.Drawing.Size(257, 459);
            this.motherTableLayoutPanel.ResumeLayout(false);
            this.mediaIconPanel.ResumeLayout(false);
            this.flowLayoutPanel1.ResumeLayout(false);
            this.myShapePanel.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel motherTableLayoutPanel;
        private System.Windows.Forms.TextBox searchBox;
        private ExtendedPanel myShapePanel;
        private ExtendedPanel mediaIconPanel;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel6;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Panel panel7;
        private System.Windows.Forms.FlowLayoutPanel myShapeFlowLayout;

    }
}
