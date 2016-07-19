using System.Drawing;
using System.Windows.Forms;

namespace PowerPointLabs.Views
{
    partial class CaptionsFormatDialogBox
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.ok = new System.Windows.Forms.Button();
            this.cancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.label3 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.boldBox = new System.Windows.Forms.CheckBox();
            this.italicBox = new System.Windows.Forms.CheckBox();
            this.fillColor = new System.Windows.Forms.Panel();
            this.fillColorDialog = new System.Windows.Forms.ColorDialog();
            this.fillColorLabel = new System.Windows.Forms.Label();
            this.fontLabel = new System.Windows.Forms.Label();
            this.fontBox = new System.Windows.Forms.ComboBox();
            this.previewText = new ReadOnlyRichTextBox();
            this.previewTextLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // ok
            // 
            this.ok.Location = new System.Drawing.Point(140, 325);
            this.ok.Name = "ok";
            this.ok.Size = new System.Drawing.Size(75, 23);
            this.ok.TabIndex = 1;
            this.ok.Text = "OK";
            this.ok.UseVisualStyleBackColor = true;
            this.ok.Click += new System.EventHandler(this.Ok_Click);
            // 
            // cancel
            // 
            this.cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancel.Location = new System.Drawing.Point(221, 325);
            this.cancel.Name = "cancel";
            this.cancel.Size = new System.Drawing.Size(75, 23);
            this.cancel.TabIndex = 2;
            this.cancel.Text = "Cancel";
            this.cancel.UseVisualStyleBackColor = true;
            this.cancel.Click += new System.EventHandler(this.Cancel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(116, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Text Size: (8-50)";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(243, 18);
            this.textBox1.MaxLength = 10;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(53, 20);
            this.textBox1.TabIndex = 4;
            this.textBox1.TabStop = false;
            this.textBox1.Validating += new System.ComponentModel.CancelEventHandler(this.TextBox1_Validating);
            // 
            // comboBox1
            // 
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(175, 44);
            this.comboBox1.Name = "alignBox";
            this.comboBox1.Size = new System.Drawing.Size(121, 21);
            this.comboBox1.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 47);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Alignment";
            // 
            // fontLabel
            // 
            this.fontLabel.AutoSize = true;
            this.fontLabel.Location = new System.Drawing.Point(12, 73);
            this.fontLabel.Name = "fontLabel";
            this.fontLabel.Size = new System.Drawing.Size(59, 13);
            this.fontLabel.TabIndex = 7;
            this.fontLabel.Text = "Font";
            // 
            // fontBox
            // 
            this.fontBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.fontBox.FormattingEnabled = true;
            this.fontBox.Location = new System.Drawing.Point(175, 70);
            this.fontBox.Name = "fontBox";
            this.fontBox.Size = new System.Drawing.Size(121, 21);
            this.fontBox.TabIndex = 8;
            this.fontBox.SelectedIndexChanged += new System.EventHandler(this.FontBox_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 98);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(75, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Text Color";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Location = new System.Drawing.Point(243, 95);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(53, 22);
            this.panel1.TabIndex = 10;
            this.panel1.Click += new System.EventHandler(this.Panel1_Click);
            // 
            // fillColorLabel
            // 
            this.fillColorLabel.AutoSize = true;
            this.fillColorLabel.Location = new System.Drawing.Point(12, 123);
            this.fillColorLabel.Name = "fillColorLabel";
            this.fillColorLabel.Size = new System.Drawing.Size(75, 13);
            this.fillColorLabel.TabIndex = 11;
            this.fillColorLabel.Text = "Background Color";
            // 
            // fillColor
            // 
            this.fillColor.BackColor = System.Drawing.Color.Black;
            this.fillColor.Location = new System.Drawing.Point(243, 120);
            this.fillColor.Name = "fillColor";
            this.fillColor.Size = new System.Drawing.Size(53, 22);
            this.fillColor.TabIndex = 12;
            this.fillColor.Click += new System.EventHandler(this.FillColor_Click);
            //
            // boldBox
            //
            this.boldBox.AutoSize = true;
            this.boldBox.Location = new System.Drawing.Point(12, 145);
            this.boldBox.Name = "boldBox";
            this.boldBox.Size = new System.Drawing.Size(148, 17);
            this.boldBox.TabIndex = 13;
            this.boldBox.Text = "Bold";
            this.boldBox.UseVisualStyleBackColor = true;
            this.boldBox.Click += new System.EventHandler(this.BoldBox_Click);
            //
            // italicBox
            //
            this.italicBox.AutoSize = true;
            this.italicBox.Location = new System.Drawing.Point(12, 170);
            this.italicBox.Name = "italixBox";
            this.italicBox.Size = new System.Drawing.Size(148, 17);
            this.italicBox.TabIndex = 14;
            this.italicBox.Text = "Italic";
            this.italicBox.UseVisualStyleBackColor = true;
            this.italicBox.Click += new System.EventHandler(this.ItalicBox_Click);
            // 
            // previewTextLabel
            // 
            this.previewTextLabel.AutoSize = true;
            this.previewTextLabel.Location = new System.Drawing.Point(12, 195);
            this.previewTextLabel.Name = "previewTextLabel";
            this.previewTextLabel.Size = new System.Drawing.Size(75, 13);
            this.previewTextLabel.TabIndex = 9;
            this.previewTextLabel.Text = "Preview:";
            // 
            // prewviewText
            // 
            this.previewText.Location = new System.Drawing.Point(12, 220);
            this.previewText.MaxLength = 10;
            this.previewText.Name = "prewviewText";
            this.previewText.Size = new System.Drawing.Size(284, 90);
            this.previewText.TabIndex = 15;
            this.previewText.TabStop = false;
            this.previewText.ReadOnly = true;
            this.previewText.ForeColor = CaptionsFormat.defaultColor;
            this.previewText.SelectAll();
            this.previewText.SelectionAlignment = HorizontalAlignment.Center;
            this.previewText.SelectedText = "ABC";
            this.previewText.ScrollBars = RichTextBoxScrollBars.None;
            this.previewText.Font = new Font(CaptionsFormat.defaultFont, 35, previewText.Font.Style);
            this.previewText.BackColor = CaptionsFormat.defaultFillColor;
            // 
            // CaptionsFormatDialogBox
            // 
            this.AcceptButton = this.ok;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.cancel;
            this.ClientSize = new System.Drawing.Size(308, 355);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.fontLabel);
            this.Controls.Add(this.fontBox);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.fillColorLabel);
            this.Controls.Add(this.fillColor);
            this.Controls.Add(this.boldBox);
            this.Controls.Add(this.italicBox);
            this.Controls.Add(this.previewTextLabel);
            this.Controls.Add(this.previewText);
            this.Controls.Add(this.cancel);
            this.Controls.Add(this.ok);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "CaptionsFormatDialogBox";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Captions Format";
            this.Load += new System.EventHandler(this.CaptionsFormatDialogBox_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public class ReadOnlyRichTextBox : RichTextBox 
        {
            private const int WM_SETFOCUS = 0x7;
            private const int WM_LBUTTONDOWN = 0x201;
            private const int WM_LBUTTONUP = 0x202;
            private const int WM_LBUTTONDBLCLK = 0x203;
            private const int WM_RBUTTONDOWN = 0x204;
            private const int WM_RBUTTONUP = 0x205;
            private const int WM_RBUTTONDBLCLK = 0x206;
            private const int WM_KEYDOWN = 0x0100;
            private const int WM_KEYUP = 0x0101;

            public ReadOnlyRichTextBox()
            {
                this.Cursor = Cursors.Arrow;  
            }

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WM_SETFOCUS
                    || m.Msg == WM_KEYDOWN
                    || m.Msg == WM_KEYUP
                    || m.Msg == WM_LBUTTONDOWN
                    || m.Msg == WM_LBUTTONUP
                    || m.Msg == WM_LBUTTONDBLCLK
                    || m.Msg == WM_RBUTTONDOWN
                    || m.Msg == WM_RBUTTONUP
                    || m.Msg == WM_RBUTTONDBLCLK)
                {
                    return;
                }
                base.WndProc(ref m);
            }
        }

        private System.Windows.Forms.Button cancel;
        private System.Windows.Forms.Button ok;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.CheckBox boldBox;
        private System.Windows.Forms.CheckBox italicBox;
        private System.Windows.Forms.Label fontLabel;
        private System.Windows.Forms.ComboBox fontBox;
        private System.Windows.Forms.Label fillColorLabel;
        private System.Windows.Forms.Panel fillColor;
        private System.Windows.Forms.ColorDialog fillColorDialog;
        private ReadOnlyRichTextBox previewText;
        private System.Windows.Forms.Label previewTextLabel;
    }
}