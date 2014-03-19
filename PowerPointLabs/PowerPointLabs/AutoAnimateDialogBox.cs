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
    public partial class AutoAnimateDialogBox : Form
    {
        public float animationDuration;
        public bool smoothAnimationChecked;
        private Ribbon1 ribbon;
        public AutoAnimateDialogBox()
        {
            InitializeComponent();
        }

        public AutoAnimateDialogBox(Ribbon1 parentRibbon, float defaultDuration, bool smoothChecked)
            : this()
        {
            ribbon = parentRibbon;
            animationDuration = defaultDuration;
            smoothAnimationChecked = smoothChecked;
        }

        private void AutoAnimateDialogBox_Load(object sender, EventArgs e)
        {
            this.textBox1.Text = animationDuration.ToString("f");
            this.checkBox1.Checked = smoothAnimationChecked;
            ToolTip ttCheckBox = new ToolTip();
            ttCheckBox.SetToolTip(checkBox1, "Use a frame-based approach for smoother resize animations.\nThis may result in larger file sizes and slower loading times for animated slides.");
            ToolTip ttTextField = new ToolTip();
            ttTextField.SetToolTip(textBox1, "The duration (in seconds) for the animations in the animation slides to be created.");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void textBox1_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            float enteredValue;
            if (float.TryParse(textBox1.Text, out enteredValue))
            {
                if (enteredValue < 0.01)
                {
                    textBox1.Text = "0.01";
                }
                else if (enteredValue > 59.0)
                {
                    textBox1.Text = "59.0";
                }
                else
                    textBox1.Text = enteredValue.ToString("f");
            }
            else if (textBox1.Text == "")
            {
                textBox1.Text = "0.01";
            }
            else
            {
                textBox1.Text = animationDuration.ToString("f"); ;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.animationDuration = float.Parse(this.textBox1.Text);
            this.smoothAnimationChecked = this.checkBox1.Checked;
            ribbon.AnimationPropertiesEdited(animationDuration, smoothAnimationChecked);
            this.Dispose();
        }
    }
}
