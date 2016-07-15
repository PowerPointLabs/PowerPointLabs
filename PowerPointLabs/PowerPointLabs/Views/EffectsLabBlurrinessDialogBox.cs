using System;
using System.Windows.Forms;

namespace PowerPointLabs.Views
{
    public partial class EffectsLabBlurrinessDialogBox : Form
    {
        public delegate void UpdateSettingsDelegate(int percentage, bool hasOverlay);
        public UpdateSettingsDelegate SettingsHandler;

        private static int previousPercentage = 90;

        public EffectsLabBlurrinessDialogBox()
        {
            InitializeComponent();
            this.numericUpDown1.Text = previousPercentage.ToString();
            this.checkBox1.Checked = EffectsLab.EffectsLabBlurSelected.HasOverlay;
        }

        private void EffectsLabBlurrinessDialogBox_Load(object sender, EventArgs e)
        {
            ToolTip ttCheckBox = new ToolTip();
            ttCheckBox.SetToolTip(checkBox1, "Insert an overlay shape to create a frosted glass effect.");
            ToolTip ttTextField = new ToolTip();
            ttTextField.SetToolTip(numericUpDown1, "The percentage of blurriness.");
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            var percentage = (int)this.numericUpDown1.Value;
            previousPercentage = percentage;
            SettingsHandler(percentage, this.checkBox1.Checked);
            this.Close();
        }
    }
}
