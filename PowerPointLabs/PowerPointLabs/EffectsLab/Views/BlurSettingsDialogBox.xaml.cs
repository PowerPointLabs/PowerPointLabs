using System.Windows;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.EffectsLab.Views
{
    /// <summary>
    /// Interaction logic for BlurSettingsDialogBox.xaml
    /// </summary>
    public partial class BlurSettingsDialogBox
    {
        public delegate void DialogConfirmedDelegate(string feature, bool isTint, int customPercentage);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

        private string _currentFeature;
        private float _blurCustomPercentage;

        public BlurSettingsDialogBox()
        {
            InitializeComponent();
        }
        
        public BlurSettingsDialogBox(string feature, bool isTint, int customPercentage)
            : this()
        {
            _currentFeature = feature;

            switch (feature)
            {
                case EffectsLabText.BlurrinessFeatureSelected:
                    Title = EffectsLabText.BlurSelectedButtonLabel + " Settings";
                    tintCheckbox.Content = EffectsLabText.SettingsTintSelectedCheckboxLabel;
                    break;
                case EffectsLabText.BlurrinessFeatureRemainder:
                    Title = EffectsLabText.BlurRemainderButtonLabel + " Settings";
                    tintCheckbox.Content = EffectsLabText.SettingsTintRemainderCheckboxLabel;
                    break;
                case EffectsLabText.BlurrinessFeatureBackground:
                    Title = EffectsLabText.BlurBackgroundButtonLabel + " Settings";
                    tintCheckbox.Content = EffectsLabText.SettingsTintBackgroundCheckboxLabel;
                    break;
                default:
                    Logger.Log(feature + " does not exist!", ActionFramework.Common.Logger.LogType.Error);
                    break;
            }

            tintCheckbox.IsChecked = isTint;
            tintCheckbox.ToolTip = EffectsLabText.SettingsTintCheckboxTooltip;

            _blurCustomPercentage = customPercentage;
            blurrinessInput.Text = (customPercentage / 100.0f).ToString("P0");
            blurrinessInput.ToolTip = EffectsLabText.SettingsBlurrinessInputTooltip;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            ValidateBlurrinessInput();
            string text = blurrinessInput.Text;
            if (text.Contains("%"))
            {
                text = text.Substring(0, text.IndexOf("%"));
            }
            int percentage = int.Parse(text);

            DialogConfirmedHandler(_currentFeature, tintCheckbox.IsChecked.GetValueOrDefault(), percentage);
            Close();
        }

        private void BlurrinessInput_LostFocus(object sender, RoutedEventArgs e)
        {
            ValidateBlurrinessInput();
        }

        private void ValidateBlurrinessInput()
        {
            float enteredValue;
            string text = blurrinessInput.Text;
            if (text.Contains("%"))
            {
                text = text.Substring(0, text.IndexOf("%"));
            }

            if (float.TryParse(text, out enteredValue) &&
                enteredValue > 0 && 
                enteredValue <= 100)
            {
                _blurCustomPercentage = enteredValue;
            }

            blurrinessInput.Text = (_blurCustomPercentage / 100.0f).ToString("P0");
        }
    }
}
