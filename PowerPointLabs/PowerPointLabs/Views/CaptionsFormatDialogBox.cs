using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using System.Collections;

namespace PowerPointLabs.Views
{
    public partial class CaptionsFormatDialogBox : Form
    {
        public delegate void UpdateSettingsDelegate(string fontName, int size, MsoTextEffectAlignment alignment, Color defaultColor, bool isBold, bool isItalic, Color defaultFillColor);
        public UpdateSettingsDelegate SettingsHandler;

        private Dictionary<String, MsoTextEffectAlignment> alignmentMapping = new Dictionary<string, MsoTextEffectAlignment>
        {
            {"Centered", MsoTextEffectAlignment.msoTextEffectAlignmentCentered},
            {"Left", MsoTextEffectAlignment.msoTextEffectAlignmentLeft},
            {"Right", MsoTextEffectAlignment.msoTextEffectAlignmentRight},
            {"Letter Justify", MsoTextEffectAlignment.msoTextEffectAlignmentLetterJustify},
            {"Stretch Justify", MsoTextEffectAlignment.msoTextEffectAlignmentStretchJustify},
            {"Word Justify", MsoTextEffectAlignment.msoTextEffectAlignmentWordJustify}
        };

        public CaptionsFormatDialogBox()
        {
            InitializeComponent();
            this.ShowInTaskbar = false;
        }

        public CaptionsFormatDialogBox(ArrayList fontList, string defaultFontName, int defaultSize, MsoTextEffectAlignment defaultAlignment, Color defaultTextColor, bool defaultBlod, bool defaultItalic, Color defaultFillColor)
            : this()
        {
            this.textBox1.Text = defaultSize.ToString();
            String[] keys = alignmentMapping.Keys.ToArray();
            this.comboBox1.Items.AddRange(keys);
            this.fontBox.DataSource = fontList;
            this.fontBox.SelectedIndex = fontList.IndexOf(CaptionsFormat.defaultFont);
            MsoTextEffectAlignment[] values = alignmentMapping.Values.ToArray();
            this.comboBox1.SelectedIndex = Array.IndexOf(values, defaultAlignment);
            panel1.BackColor = defaultTextColor;
            this.boldBox.Checked = defaultBlod;
            this.italicBox.Checked = defaultItalic;
            fillColor.BackColor = defaultFillColor;
        }

        private void CaptionsFormatDialogBox_Load(object sender, EventArgs e)
        {
            ToolTip ttComboBox = new ToolTip();
            ttComboBox.SetToolTip(comboBox1, "The alignment of the Captions.");

            ToolTip fontComboBox = new ToolTip();
            fontComboBox.SetToolTip(fontBox, "The font of the text.");
        }

        private void TextBox1_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            string text = textBox1.Text;
            int thisSize = Int32.Parse(text);
            int max = 50;
            int min = 8;
            if (thisSize >= max)
            {
                textBox1.Text = max.ToString();
            }
            if (thisSize <= min)
            {
                textBox1.Text = min.ToString();
            }
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Ok_Click(object sender, EventArgs e)
        {
            string text = textBox1.Text;

            SettingsHandler(fontBox.Text, Int32.Parse(text), alignmentMapping[(String)this.comboBox1.SelectedItem], panel1.BackColor, this.boldBox.Checked, this.italicBox.Checked, fillColor.BackColor);
            if (Ribbon1.HaveCaptions)
            {
                NotesToCaptions.EmbedCaptionsOnSelectedSlides();
            }
            Close();
        }

        private void Panel1_Click(object sender, EventArgs e)
        {
            colorDialog1.Color = panel1.BackColor;
            colorDialog1.FullOpen = true;
            DialogResult result = colorDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                panel1.BackColor = colorDialog1.Color;
            }
            UpdatePreviewText();
        }

        private void FillColor_Click(object sender, EventArgs e)
        {
            fillColorDialog.Color = fillColor.BackColor;
            fillColorDialog.FullOpen = true;
            DialogResult result = fillColorDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                fillColor.BackColor = fillColorDialog.Color;
            }
            UpdatePreviewText();
        }

        private void FontBox_Click(object sender, EventArgs e)
        {
            UpdatePreviewText();
        }

        private void BoldBox_Click(object sender, EventArgs e)
        {
            UpdatePreviewText();
        }

        private void ItalicBox_Click(object sender, EventArgs e)
        {
            UpdatePreviewText();
        }

        private void UpdatePreviewText()
        {
            this.prewviewText.BackColor = this.fillColor.BackColor;
            this.prewviewText.ForeColor = this.panel1.BackColor;
            if (this.boldBox.Checked && this.italicBox.Checked)
            {
                this.prewviewText.Font = new Font(this.fontBox.Text, this.prewviewText.Font.Size,
                                                  FontStyle.Bold | FontStyle.Italic);
            }
            else if (this.boldBox.Checked)
            {
                this.prewviewText.Font = new Font(this.fontBox.Text, this.prewviewText.Font.Size,
                                                  FontStyle.Bold);
            }
            else if (this.italicBox.Checked)
            {
                this.prewviewText.Font = new Font(this.fontBox.Text, this.prewviewText.Font.Size,
                                                  FontStyle.Italic);
            }
            else
            {
                this.prewviewText.Font = new Font(this.fontBox.Text, this.prewviewText.Font.Size, 
                                                  FontStyle.Regular);
            }          
        }
    }
}
