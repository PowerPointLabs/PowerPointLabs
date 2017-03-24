using System.Windows;

namespace PowerPointLabs.CropLab
{
    /// <summary>
    /// Interaction logic for CropLabSettingsDialog.xaml
    /// </summary>
    public partial class CropLabSettingsDialog
    {
        private AnchorPosition selectedAnchor;

        public CropLabSettingsDialog()
        {
            selectedAnchor = CropLabSettings.AnchorPosition;
            InitializeComponent();
        }

        public AnchorPosition SelectedAnchor
        {
            get { return selectedAnchor; }
            set { selectedAnchor = value; }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            CropLabSettings.AnchorPosition = selectedAnchor;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
