using System.Windows;

namespace PowerPointLabs.PositionsLab.Views
{
    /// <summary>
    /// Interaction logic for AlignSettingsDialog.xaml
    /// </summary>
    public partial class AlignSettingsDialog
    {
        public delegate void DialogConfirmedDelegate(PositionsLabSettings.AlignReferenceObject alignReference);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

        public AlignSettingsDialog()
        {
            InitializeComponent();
        }

        public AlignSettingsDialog(PositionsLabSettings.AlignReferenceObject alignReference)
            : this()
        {
            switch (alignReference)
            {
                case PositionsLabSettings.AlignReferenceObject.Slide:
                    alignToSlideButton.IsChecked = true;
                    break;
                case PositionsLabSettings.AlignReferenceObject.SelectedShape:
                    alignToShapeButton.IsChecked = true;
                    break;
                case PositionsLabSettings.AlignReferenceObject.PowerpointDefaults:
                    alignPowerpointDefaultsButton.IsChecked = true;
                    break;
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (alignToSlideButton.IsChecked.GetValueOrDefault())
            {
                DialogConfirmedHandler(PositionsLabSettings.AlignReferenceObject.Slide);
            }
            else if (alignToShapeButton.IsChecked.GetValueOrDefault())
            {
                DialogConfirmedHandler(PositionsLabSettings.AlignReferenceObject.SelectedShape);
            }
            else
            {
                DialogConfirmedHandler(PositionsLabSettings.AlignReferenceObject.PowerpointDefaults);
            }
            Close();
        }
    }
}
