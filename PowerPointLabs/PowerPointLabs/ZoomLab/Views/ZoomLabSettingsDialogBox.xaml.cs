using System.Windows;

using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ZoomLab.Views
{
    /// <summary>
    /// Interaction logic for ZoomLabSettingsDialogBox.xaml
    /// </summary>
    public partial class ZoomLabSettingsDialogBox
    {
        public delegate void DialogConfirmedDelegate(bool slideBackgroundChecked, bool multiSlideChecked);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

        public ZoomLabSettingsDialogBox()
        {
            InitializeComponent();
        }
        
        public ZoomLabSettingsDialogBox(bool backgroundChecked, bool multiChecked)
            : this()
        {
            slideBackgroundCheckbox.IsChecked = backgroundChecked;
            slideBackgroundCheckbox.ToolTip = ZoomLabText.SettingsSlideBackgroundCheckboxTooltip;

            separateSlidesCheckbox.IsChecked = multiChecked;
            separateSlidesCheckbox.ToolTip = ZoomLabText.SettingsSeparateSlidesCheckboxTooltip;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogConfirmedHandler(slideBackgroundCheckbox.IsChecked.GetValueOrDefault(), 
                            separateSlidesCheckbox.IsChecked.GetValueOrDefault());
            Close();
        }
    }
}
