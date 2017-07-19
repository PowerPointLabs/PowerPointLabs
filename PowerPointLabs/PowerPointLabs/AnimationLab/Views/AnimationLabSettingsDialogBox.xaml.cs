using System;
using System.Windows;

using Microsoft.Office.Interop.PowerPoint;

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
            durationInput.ToolTip = TextCollection.AnimationLabSettingsDurationInputTooltip;
            durationInput.SelectAll();

            smoothAnimationCheckbox.IsChecked = smoothChecked;
            smoothAnimationCheckbox.ToolTip = TextCollection.AnimationLabSettingsSmoothAnimationCheckboxTooltip;

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
