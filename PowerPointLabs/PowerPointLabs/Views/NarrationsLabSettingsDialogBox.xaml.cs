using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

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
        
        public NarrationsLabSettingsDialogBox(int selectedVoiceIndex, List<string> voices, bool isPreviewChecked)
            : this()
        {
            voiceSelectionInput.ItemsSource = voices;
            voiceSelectionInput.ToolTip = TextCollection.NarrationsLabSettingsVoiceSelectionInputTooltip;
            voiceSelectionInput.Content = voices[selectedVoiceIndex];

            previewCheckbox.IsChecked = isPreviewChecked;
            previewCheckbox.ToolTip = TextCollection.NarrationsLabSettingsPreviewCheckboxTooltip;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            SettingsHandler(voiceSelectionInput.Content.ToString(), previewCheckbox.IsChecked.GetValueOrDefault());
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        void VoiceSelectionInput_Item_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left && voiceSelectionInput.IsExpanded)
            {
                string value = ((TextBlock)e.Source).Text;
                voiceSelectionInput.Content = value;
            }
        }
    }
}
