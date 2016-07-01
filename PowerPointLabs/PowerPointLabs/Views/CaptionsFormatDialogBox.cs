using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;

namespace PowerPointLabs.Views
{
    public partial class CaptionsFormatDialogBox : Form
    {
        public delegate void UpdateSettingsDelegate(int size, MsoTextEffectAlignment alignment, Color defaultColor, bool isBold, bool isItalic);
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

        public CaptionsFormatDialogBox(int defaultSize, MsoTextEffectAlignment defaultAlignment, Color defaultColor, bool defaultBlod, bool defaultItalic)
            : this()
        {
            this.textBox1.Text = defaultSize.ToString();
            String[] keys = alignmentMapping.Keys.ToArray();
            this.comboBox1.Items.AddRange(keys);
            MsoTextEffectAlignment[] values = alignmentMapping.Values.ToArray();
            this.comboBox1.SelectedIndex = Array.IndexOf(values, defaultAlignment);
            panel1.BackColor = defaultColor;
            this.boldBox.Checked = defaultBlod;
            this.italicBox.Checked = defaultItalic;
        }

        private void CaptionsFormatDialogBox_Load(object sender, EventArgs e)
        {
            ToolTip ttComboBox = new ToolTip();
            ttComboBox.SetToolTip(comboBox1, "The alignment of the Captions.");
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

            SettingsHandler(Int32.Parse(text), alignmentMapping[(String)this.comboBox1.SelectedItem], panel1.BackColor, this.boldBox.Checked, this.italicBox.Checked);
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
        }
    }
}
