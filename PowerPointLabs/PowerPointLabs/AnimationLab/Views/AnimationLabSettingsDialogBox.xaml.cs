using System.Collections;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

using PowerPointLabs.TextCollection;

namespace PowerPointLabs.AnimationLab.Views
{
    /// <summary>
    /// Interaction logic for AnimationLabSettingsDialogBox.xaml
    /// </summary>
    public partial class AnimationLabSettingsDialogBox
    {
        public delegate void DialogConfirmedDelegate(float animationDuration, bool smoothAnimationChecked);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

        private float lastDuration;

        public AnimationLabSettingsDialogBox()
        {
            InitializeComponent();
        }
        
        public AnimationLabSettingsDialogBox(float defaultDuration, bool smoothChecked)
            : this()
        {
            durationInput.Text = defaultDuration.ToString("f");
            durationInput.ToolTip = AnimationLabText.SettingsDurationInputTooltip;
            durationInput.SelectAll();

            smoothAnimationCheckbox.IsChecked = smoothChecked;
            smoothAnimationCheckbox.ToolTip = AnimationLabText.SettingsSmoothAnimationCheckboxTooltip;

            lastDuration = defaultDuration;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            ValidateDurationInput();
            DialogConfirmedHandler(float.Parse(durationInput.Text), smoothAnimationCheckbox.IsChecked.GetValueOrDefault());
            Close();
        }

        private void DurationInput_LostFocus(object sender, RoutedEventArgs e)
        {
            ValidateDurationInput();
        }

        private void ValidateDurationInput()
        {
            float enteredValue;
            if (float.TryParse(durationInput.Text, out enteredValue))
            {
                if (enteredValue < 0.01)
                {
                    enteredValue = 0.01f;
                }
                else if (enteredValue > 59.0)
                {
                    enteredValue = 59.0f;
                }
            }
            else
            {
                enteredValue = lastDuration;
            }
            durationInput.Text = enteredValue.ToString("f");
            lastDuration = enteredValue;
        }
    }
}
