using System.Collections.Generic;
using System.Windows;

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
            slideBackgroundCheckbox.ToolTip = TextCollection1.ZoomLabSettingsSlideBackgroundCheckboxTooltip;

            separateSlidesCheckbox.IsChecked = multiChecked;
            separateSlidesCheckbox.ToolTip = TextCollection1.ZoomLabSettingsSeparateSlidesCheckboxTooltip;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogConfirmedHandler(slideBackgroundCheckbox.IsChecked.GetValueOrDefault(), 
                            separateSlidesCheckbox.IsChecked.GetValueOrDefault());
            Close();
        }
    }
}
