using System.Windows;

using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.EffectsLab.Views
{
    /// <summary>
    /// Interaction logic for EffectsLabBlurDialogBox.xaml
    /// </summary>
    public partial class EffectsLabBlurDialogBox
    {
        public delegate void UpdateSettingsDelegate(int percentage, bool isTinted);
        public UpdateSettingsDelegate SettingsHandler { get; set; }

        private string currentFeature;
        private float lastBlurriness;

        public EffectsLabBlurDialogBox()
        {
            InitializeComponent();
        }
        
        public EffectsLabBlurDialogBox(string feature)
            : this()
        {
            currentFeature = feature;
            string properFeatureName = "Effects Lab";

            switch (feature)
            {
                case TextCollection.EffectsLabBlurrinessFeatureSelected:
                    properFeatureName = TextCollection.EffectsLabBlurSelectedButtonLabel;
                    lastBlurriness = EffectsLabBlurSelected.CustomPercentageSelected;
                    tintCheckbox.IsChecked = EffectsLabBlurSelected.IsTintSelected;
                    break;
                case TextCollection.EffectsLabBlurrinessFeatureRemainder:
                    properFeatureName = TextCollection.EffectsLabBlurRemainderButtonLabel;
                    lastBlurriness = EffectsLabBlurSelected.CustomPercentageRemainder;
                    tintCheckbox.IsChecked = EffectsLabBlurSelected.IsTintRemainder;
                    break;
                case TextCollection.EffectsLabBlurrinessFeatureBackground:
                    properFeatureName = TextCollection.EffectsLabBlurBackgroundButtonLabel;
                    lastBlurriness = EffectsLabBlurSelected.CustomPercentageBackground;
                    tintCheckbox.IsChecked = EffectsLabBlurSelected.IsTintBackground;
                    break;
                default:
                    Logger.Log(feature + " does not exist!", ActionFramework.Common.Logger.LogType.Error);
                    break;
            }

            Title = properFeatureName + " Settings";

            tintCheckbox.Content = properFeatureName.Replace("Blur", "Tint");
            tintCheckbox.ToolTip = "Adds a tinted effect to your blur.";

            blurrinessInput.Text = (lastBlurriness / 100.0f).ToString("P0");
            blurrinessInput.ToolTip = "The level of blurriness.";
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            BlurrinessInput_LostFocus(null, null);
            string text = blurrinessInput.Text;
            if (text.Contains("%"))
            {
                text = text.Substring(0, text.IndexOf("%"));
            }
            int percentage = int.Parse(text);

            SettingsHandler(percentage, tintCheckbox.IsChecked.GetValueOrDefault());
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void BlurrinessInput_LostFocus(object sender, RoutedEventArgs e)
        {
            float enteredValue;
            string text = blurrinessInput.Text;
            if (text.Contains("%"))
            {
                text = text.Substring(0, text.IndexOf("%"));
            }

            if (float.TryParse(text, out enteredValue))
            {
                if (enteredValue > 0 && enteredValue <= 100)
                {
                    lastBlurriness = enteredValue / 100;
                }
                blurrinessInput.Text = lastBlurriness.ToString("P0");
            }
            else
            {
                blurrinessInput.Text = lastBlurriness.ToString("P0");
            }
        }
    }
}
