using System.Windows;

namespace PowerPointLabs.CropLab.Views
{
    /// <summary>
    /// Interaction logic for CropLabSettingsDialogBox.xaml
    /// </summary>
    public partial class CropLabSettingsDialogBox
    {
        public CropLabSettingsDialogBox()
        {
            SelectedAnchor = CropLabSettings.AnchorPosition;
            InitializeComponent();
        }

        public AnchorPosition SelectedAnchor { get; set; }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            CropLabSettings.AnchorPosition = SelectedAnchor;
            Close();
        }
    }
}
