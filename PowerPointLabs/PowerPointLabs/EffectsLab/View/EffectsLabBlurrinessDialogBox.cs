using System;
using System.Windows.Forms;

namespace PowerPointLabs.EffectsLab.View
{
    public partial class EffectsLabBlurrinessDialogBox : Form
    {
        public delegate void UpdateSettingsDelegate(int percentage, bool hasOverlay);
        public UpdateSettingsDelegate SettingsHandler;

        private static int previousPercentageSelected = 90;
        private static int previousPercentageRemainder = 90;
        private static int previousPercentageBackground = 90;

        private string _feature;

        public EffectsLabBlurrinessDialogBox(string feature)
        {
            InitializeComponent();

            _feature = feature;

            var startIndex = feature.IndexOf("Blur") + 4;
            var tintFeatureText = feature.Substring(startIndex, feature.Length - startIndex);
            this.checkBox1.Text += tintFeatureText;

            switch (feature)
            {
                case TextCollection.EffectsLabBlurrinessFeatureSelected:
                    this.numericUpDown1.Value = previousPercentageSelected;
                    this.checkBox1.Checked = EffectsLabBlurSelected.IsTintSelected;
                    break;
                case TextCollection.EffectsLabBlurrinessFeatureRemainder:
                    this.numericUpDown1.Value = previousPercentageRemainder;
                    this.checkBox1.Checked = EffectsLabBlurSelected.IsTintRemainder;
                    break;
                case TextCollection.EffectsLabBlurrinessFeatureBackground:
                    this.numericUpDown1.Value = previousPercentageBackground;
                    this.checkBox1.Checked = EffectsLabBlurSelected.IsTintBackground;
                    break;
                default:
                    throw new System.Exception("Invalid feature");
            }
        }

        private void EffectsLabBlurrinessDialogBox_Load(object sender, EventArgs e)
        {
            ToolTip ttCheckBox = new ToolTip();
            ttCheckBox.SetToolTip(checkBox1, "Insert an overlay shape to create a tinted effect.");
            ToolTip ttTextField = new ToolTip();
            ttTextField.SetToolTip(numericUpDown1, "The level of blurriness.");
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            var percentage = (int)this.numericUpDown1.Value;

            switch (_feature)
            {
                case TextCollection.EffectsLabBlurrinessFeatureSelected:
                    previousPercentageSelected = percentage;
                    break;
                case TextCollection.EffectsLabBlurrinessFeatureRemainder:
                    previousPercentageRemainder = percentage;
                    break;
                case TextCollection.EffectsLabBlurrinessFeatureBackground:
                    previousPercentageBackground = percentage;
                    break;
                default:
                    throw new System.Exception("Invalid feature");
            }


            SettingsHandler(percentage, this.checkBox1.Checked);
            this.Close();
        }
    }
}
