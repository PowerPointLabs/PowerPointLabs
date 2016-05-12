using System.Windows;

namespace PowerPointLabs.PositionsLab
{
    /// <summary>
    /// Interaction logic for AlignSettingsDialog.xaml
    /// </summary>
    public partial class AlignSettingsDialog
    {
        //Flag to trigger
        public bool IsOpen { get; set; }

        public AlignSettingsDialog()
        {
            IsOpen = true;
            InitializeComponent();
        }

        private void AlignToSlideButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.AlignReference == PositionsLabMain.AlignReferenceObject.Slide)
            {
                alignToSlideButton.IsChecked = true;
            }
        }

        private void AlignToShapeButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.AlignReference == PositionsLabMain.AlignReferenceObject.SelectedShape)
            {
                alignToShapeButton.IsChecked = true;
            }
        }

        private void PowerpointAlignButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.AlignReference == PositionsLabMain.AlignReferenceObject.PowerpointDefaults)
            {
                alignPowerpointDefaultsButton.IsChecked = true;
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (alignToShapeButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.AlignReferToShape();
            }
            else if (alignToSlideButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.AlignReferToSlide();
            }
            else
            {
                PositionsLabMain.AlignReferToPowerpointDefaults();
            }

            IsOpen = false;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            IsOpen = false;
            Close();
        }

        private void AlignSettingsDialong_Closed(object sender, System.EventArgs e)
        {
            IsOpen = false;
        }
    }
}
