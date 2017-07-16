using System.Windows;

namespace PowerPointLabs.CropLab.Views
{
    /// <summary>
    /// Interaction logic for CropLabSettingsDialogBox.xaml
    /// </summary>
    public partial class CropLabSettingsDialogBox
    {
        public delegate void UpdateSettingsDelegate(AnchorPosition anchorPosition);
        public UpdateSettingsDelegate SettingsHandler { get; set; }

        public CropLabSettingsDialogBox()
        {
            // Special case: Anchor point must be set before InitializeComponent
            //InitializeComponent();
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
            SettingsHandler(SelectedAnchor);
            Close();
        }
    }
}
