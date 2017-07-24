using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

using PowerPointLabs.TextCollection;

namespace PowerPointLabs.NarrationsLab.Views
{
    /// <summary>
    /// Interaction logic for NarrationsLabSettingsDialogBox.xaml
    /// </summary>
    public partial class NarrationsLabSettingsDialogBox
    {
        public delegate void DialogConfirmedDelegate(string voiceName, bool preview);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

        public NarrationsLabSettingsDialogBox()
        {
            InitializeComponent();
        }
        
        public NarrationsLabSettingsDialogBox(int selectedVoiceIndex, List<string> voices, bool isPreviewChecked)
            : this()
        {
            voiceSelectionInput.ItemsSource = voices;
            voiceSelectionInput.ToolTip = NarrationsLabText.SettingsVoiceSelectionInputTooltip;
            voiceSelectionInput.Content = voices[selectedVoiceIndex];

            previewCheckbox.IsChecked = isPreviewChecked;
            previewCheckbox.ToolTip = NarrationsLabText.SettingsPreviewCheckboxTooltip;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogConfirmedHandler(voiceSelectionInput.Content.ToString(), previewCheckbox.IsChecked.GetValueOrDefault());
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
