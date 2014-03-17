using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PowerPointLabs
{
    public partial class SpotlightDialogBox : Form
    {
        public float spotlightTransparency;
        public float softEdge;
        public Dictionary<String, float> softEdgesMapping = new Dictionary<string, float>
        {
            {"None", 0},
            {"1 Point", 1},
            {"2.5 Points", 2.5f},
            {"5 Points", 5},
            {"10 Points", 10},
            {"25 Points", 25},
            {"50 Points", 50}
        };
        private Ribbon1 ribbon;
        public SpotlightDialogBox()
        {
            InitializeComponent();
        }

        public SpotlightDialogBox(Ribbon1 parentRibbon, float defaultTransparency, float defaultSoftEdge)
            : this()
        {
            ribbon = parentRibbon;
            spotlightTransparency = defaultTransparency;
            softEdge = defaultSoftEdge;
            String[] keys = softEdgesMapping.Keys.ToArray();
            this.comboBox1.Items.AddRange(keys);
        }

        private void SpotlightDialogBox_Load(object sender, EventArgs e)
        {
            this.textBox1.Text = spotlightTransparency.ToString("P0");
            float[] values = softEdgesMapping.Values.ToArray();
            this.comboBox1.SelectedIndex = Array.IndexOf(values, softEdge);

            ToolTip ttComboBox = new ToolTip();
            ttComboBox.SetToolTip(comboBox1, "The softness of the edges of the spotlight effect to be created.");
            ToolTip ttTextField = new ToolTip();
            ttTextField.SetToolTip(textBox1, "The transparency level of the spotlight effect to be created.");
        }

        private void textBox1_Validating(object sender, System.ComponentModel.CancelEventArgs e)
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
                    spotlightTransparency = enteredValue / 100;
                }
                textBox1.Text = spotlightTransparency.ToString("P0");
            }
            else
            {
                textBox1.Text = spotlightTransparency.ToString("P0"); ;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string text = textBox1.Text;
            if (text.Contains('%'))
            {
                text = text.Substring(0, text.IndexOf('%'));
            }

            this.spotlightTransparency = float.Parse(text) / 100;
            this.softEdge = softEdgesMapping[(String)this.comboBox1.SelectedItem];
            ribbon.SpotlightPropertiesEdited(spotlightTransparency, softEdge);
            this.Dispose();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }
    }
}
