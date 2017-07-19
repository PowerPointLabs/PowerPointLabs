using System.Collections.Generic;
using System.Windows;

namespace PowerPointLabs.Views
{
    /// <summary>
    /// Interaction logic for ZoomLabSettingsDialogBox.xaml
    /// </summary>
    public partial class ZoomLabSettingsDialogBox
    {
        public delegate void UpdateSettingsDelegate(bool slideBackgroundChecked, bool multiSlideChecked);
        public UpdateSettingsDelegate SettingsHandler { get; set; }

        public ZoomLabSettingsDialogBox()
        {
            InitializeComponent();
        }
        
        public ZoomLabSettingsDialogBox(bool backgroundChecked, bool multiChecked)
            : this()
        {
            slideBackgroundCheckbox.IsChecked = backgroundChecked;
            slideBackgroundCheckbox.ToolTip = TextCollection.ZoomLabSettingsSlideBackgroundCheckboxTooltip;

            separateSlidesCheckbox.IsChecked = multiChecked;
            separateSlidesCheckbox.ToolTip = TextCollection.ZoomLabSettingsSeparateSlidesCheckboxTooltip;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            SettingsHandler(slideBackgroundCheckbox.IsChecked.GetValueOrDefault(), 
                            separateSlidesCheckbox.IsChecked.GetValueOrDefault());
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
