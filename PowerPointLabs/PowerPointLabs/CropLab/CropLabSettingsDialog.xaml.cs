using System.Windows;

namespace PowerPointLabs.CropLab
{
    /// <summary>
    /// Interaction logic for CropLabSettingsDialog.xaml
    /// </summary>
    public partial class CropLabSettingsDialog
    {
        public CropLabSettingsDialog()
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
