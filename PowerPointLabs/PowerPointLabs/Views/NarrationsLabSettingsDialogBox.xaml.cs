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
        public UpdateSettingsDelegate SettingsHandler { get; set; }

        public NarrationsLabSettingsDialogBox()
        {
            InitializeComponent();
        }
        
        public NarrationsLabSettingsDialogBox(int selectedVoice, List<string> voices, bool isPreviewChecked)
            : this()
        {
            voiceSelectionInput.ItemsSource = voices;
            voiceSelectionInput.SelectedIndex = selectedVoice;
            voiceSelectionInput.ToolTip = TextCollection.NarrationsLabSettingsVoiceSelectionInputTooltip;

            previewCheckbox.IsChecked = isPreviewChecked;
            previewCheckbox.ToolTip = TextCollection.NarrationsLabSettingsPreviewCheckboxTooltip;
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
