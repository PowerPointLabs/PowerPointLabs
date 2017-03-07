using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace PowerPointLabs.Views
{
    public partial class SpotlightDialogBox : Form
    {
        public delegate void UpdateSettingsDelegate(float spotlightTransparency, float softEdge, Color newColor);
        public UpdateSettingsDelegate SettingsHandler;

        private Dictionary<String, float> softEdgesMapping = new Dictionary<string, float>
        {
            {"None", 0},
            {"1 Point", 1},
            {"2.5 Points", 2.5f},
            {"5 Points", 5},
            {"10 Points", 10},
            {"25 Points", 25},
            {"50 Points", 50}
        };
        private float lastTransparency;

        public SpotlightDialogBox()
        {
            InitializeComponent();
        }

        public SpotlightDialogBox(float defaultTransparency, float defaultSoftEdge, Color defaultColor)
            : this()
        {
            this.textBox1.Text = defaultTransparency.ToString("P0");
            lastTransparency = defaultTransparency;

            String[] keys = softEdgesMapping.Keys.ToArray();
            this.comboBox1.Items.AddRange(keys);
            float[] values = softEdgesMapping.Values.ToArray();
            this.comboBox1.SelectedIndex = Array.IndexOf(values, defaultSoftEdge);
            panel1.BackColor = defaultColor;
        }

        private void SpotlightDialogBox_Load(object sender, EventArgs e)
        {
            ToolTip ttComboBox = new ToolTip();
            ttComboBox.SetToolTip(comboBox1, "The softness of the edges of the spotlight effect to be created.");
            ToolTip ttTextField = new ToolTip();
            ttTextField.SetToolTip(textBox1, "The transparency level of the spotlight effect to be created.");
        }

        private void TextBox1_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            float enteredValue;
            string text = textBox1.Text;
            if (text.Contains('%'))
            {
                text = text.Substring(0, text.IndexOf('%'));
            }

            if (float.TryParse(text, out enteredValue))
            {
                if (enteredValue > 0 && enteredValue <= 100)
                {
                    lastTransparency = enteredValue / 100;
                }
                textBox1.Text = lastTransparency.ToString("P0");
            }
            else
            {
                textBox1.Text = lastTransparency.ToString("P0");
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            string text = textBox1.Text;
            if (text.Contains('%'))
            {
                text = text.Substring(0, text.IndexOf('%'));
            }

            SettingsHandler(float.Parse(text) / 100, softEdgesMapping[(String)this.comboBox1.SelectedItem], panel1.BackColor);
            this.Close();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            this.Close();
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
