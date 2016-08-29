using System.Windows;

namespace PowerPointLabs.PositionsLab
{
    /// <summary>
    /// Interaction logic for DistributeSettingsDialog.xaml
    /// </summary>
    public partial class DistributeSettingsDialog
    {
        //Flag to trigger
        public bool IsOpen { get; set; }

        public DistributeSettingsDialog()
        {
            IsOpen = true;
            InitializeComponent();
        }

        #region On-Load Settings
        private void DistributeToSlideButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.DistributeReference == PositionsLabMain.DistributeReferenceObject.Slide)
            {
                distributeToSlideButton.IsChecked = true;
            }
        }

        private void DistributeToFirstShapeButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.DistributeReference == PositionsLabMain.DistributeReferenceObject.FirstShape)
            {
                distributeToFirstShapeButton.IsChecked = true;
            }
        }

        private void DistributeToFirstTwoShapesButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.DistributeReference == PositionsLabMain.DistributeReferenceObject.FirstTwoShapes)
            {
                distributeToFirstTwoShapesButton.IsChecked = true;
            }
        }

        private void DistributeExtremeShapesButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.DistributeReference == PositionsLabMain.DistributeReferenceObject.ExtremeShapes)
            {
                distributeToExtremeShapesButton.IsChecked = true;
            }
        }

        private void DistributeByBoundariesButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.DistributeSpaceReference == PositionsLabMain.DistributeSpaceReferenceObject.ObjectBoundary)
            {
                distributeByBoundariesButton.IsChecked = true;
            }
        }

        private void DistributeByShapeCenterButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.DistributeSpaceReference == PositionsLabMain.DistributeSpaceReferenceObject.ObjectCenter)
            {
                distributeByShapeCenterButton.IsChecked = true;
            }
        }
        #endregion

        #region Button actions
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            // Checks for boundary reference
            if (distributeToSlideButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.DistributeReferToSlide();
            }

            if (distributeToFirstShapeButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.DistributeReferToFirstShape();
            }

            if (distributeToFirstTwoShapesButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.DistributeReferToFirstTwoShapes();
            }

            if (distributeToExtremeShapesButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.DistributeReferToExtremeShapes();
            }

            // Checks for space calculation reference
            if (distributeByBoundariesButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.DistributeSpaceByBoundaries();
            }

            if (distributeByShapeCenterButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.DistributeSpaceByCenter();
            }

            IsOpen = false;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            IsOpen = false;
            Close();
        }

        private void DistributeSettingsDialong_Closed(object sender, System.EventArgs e)
        {
            IsOpen = false;
        }
        #endregion
    }
}
