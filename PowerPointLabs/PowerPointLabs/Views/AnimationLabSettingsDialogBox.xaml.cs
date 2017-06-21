using System;
using System.Windows;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Views
{
    /// <summary>
    /// Interaction logic for AnimationLabSettingsDialogBox.xaml
    /// </summary>
    public partial class AnimationLabSettingsDialogBox
    {
        public delegate void UpdateSettingsDelegate(float animationDuration, bool smoothAnimationChecked);
        public UpdateSettingsDelegate SettingsHandler { get; set; }

        private float lastDuration;

        public AnimationLabSettingsDialogBox()
        {
            InitializeComponent();
        }
        
        public AnimationLabSettingsDialogBox(float defaultDuration, bool smoothChecked)
            : this()
        {
            durationInput.Text = defaultDuration.ToString("f");
            durationInput.ToolTip = 
                "The duration (in seconds) for the animations in the animation slides to be created.";
            durationInput.SelectAll();

            smoothAnimationCheckbox.IsChecked = smoothChecked;
            smoothAnimationCheckbox.ToolTip = 
                "Use a frame-based approach for smoother resize animations.\n" +
                "This may result in larger file sizes and slower loading times for animated slides.";

            lastDuration = defaultDuration;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DurationInput_LostFocus(null, null);
            SettingsHandler(float.Parse(durationInput.Text), smoothAnimationCheckbox.IsChecked.GetValueOrDefault());
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void DurationInput_LostFocus(object sender, RoutedEventArgs e)
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
