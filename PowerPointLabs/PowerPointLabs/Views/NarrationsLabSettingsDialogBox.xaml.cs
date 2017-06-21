using System;
using System.Collections.Generic;
using System.Windows;

namespace PowerPointLabs.Views
{
    /// <summary>
    /// Interaction logic for NarrationsLabSettingsDialogBox.xaml
    /// </summary>
    public partial class NarrationsLabSettingsDialogBox
    {
        public delegate void UpdateSettingsDelegate(string voiceName, bool preview);
        public UpdateSettingsDelegate SettingsHandler;

        public NarrationsLabSettingsDialogBox()
        {
            InitializeComponent();
        }
        
        public NarrationsLabSettingsDialogBox(int selectedVoice, List<string> voices, bool isPreviewChecked)
            : this()
        {
            voiceSelectionInput.ItemsSource = voices;
            voiceSelectionInput.SelectedIndex = selectedVoice;
            voiceSelectionInput.ToolTip =
                "The voice to be used when generating synthesized audio. " +
                "Use [Voice] tags to specify a different voice for a particular section of text.";

            previewCheckbox.IsChecked = isPreviewChecked;
            previewCheckbox.ToolTip =
                "If checked, the current slide's audio and animations will play after the Add Audio button is clicked.";
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            SettingsHandler(voiceSelectionInput.SelectedItem.ToString(), previewCheckbox.IsChecked.GetValueOrDefault());
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
