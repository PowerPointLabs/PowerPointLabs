using System.Windows;

namespace PowerPointLabs.CropLab.Views
{
    /// <summary>
    /// Interaction logic for CropLabSettingsDialogBox.xaml
    /// </summary>
    public partial class CropLabSettingsDialogBox
    {
        public delegate void DialogConfirmedDelegate(AnchorPosition anchorPosition);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

        public CropLabSettingsDialogBox()
        {
        }

        public CropLabSettingsDialogBox(AnchorPosition anchorPosition)
            : this()
        {
            SelectedAnchor = anchorPosition;
            InitializeComponent();
        }

        public AnchorPosition SelectedAnchor { get; set; }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogConfirmedHandler(SelectedAnchor);
            Close();
        }
    }
}
