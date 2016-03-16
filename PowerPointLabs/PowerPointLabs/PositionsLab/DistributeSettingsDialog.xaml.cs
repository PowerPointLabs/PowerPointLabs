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
        private void MarginTopInput_Load(object sender, RoutedEventArgs e)
        {
            marginTopInput.Value = PositionsLabMain.MarginTop;
        }

        private void MarginBottomInput_Load(object sender, RoutedEventArgs e)
        {
            marginBottomInput.Value = PositionsLabMain.MarginBottom;
        }

        private void MarginLeftInput_Load(object sender, RoutedEventArgs e)
        {
            marginLeftInput.Value = PositionsLabMain.MarginLeft;
        }

        private void MarginRightInput_Load(object sender, RoutedEventArgs e)
        {
            marginRightInput.Value = PositionsLabMain.MarginRight;
        }

        private void AlignLeftButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.DistributeGridAlignment == PositionsLabMain.GridAlignment.AlignLeft)
            {
                alignLeftButton.IsChecked = true;
            }
        }

        private void AlignCenterButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.DistributeGridAlignment == PositionsLabMain.GridAlignment.AlignCenter)
            {
                alignCenterButton.IsChecked = true;
            }
        }

        private void AlignRightButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.DistributeGridAlignment == PositionsLabMain.GridAlignment.AlignRight)
            {
                alignRightButton.IsChecked = true;
            }
        }

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
            var marginTopValue = marginTopInput.Value;
            var marginBottomValue = marginBottomInput.Value;
            var marginLeftValue = marginLeftInput.Value;
            var marginRightValue = marginRightInput.Value;

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
                PositionsLabMain.DistributeRefertoFirstTwoShapes();
            }

            if (distributeByBoundariesButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.DistributeSpaceByBoundaries();
            }

            if (distributeByShapeCenterButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.DistributeSpaceByCenter();
            }


            if (!marginTopValue.HasValue || marginTopValue.GetValueOrDefault() < 0 ||
                !marginBottomValue.HasValue || marginBottomValue.GetValueOrDefault() < 0 ||
                !marginLeftValue.HasValue || marginLeftValue.GetValueOrDefault() < 0 ||
                !marginRightValue.HasValue || marginRightValue.GetValueOrDefault() < 0)
            {
                // TODO: Notify the user that not successfully changed
                return;
            }

            PositionsLabMain.SetDistributeMarginTop((float)marginTopValue);
            PositionsLabMain.SetDistributeMarginBottom((float)marginBottomValue);
            PositionsLabMain.SetDistributeMarginLeft((float)marginLeftValue);
            PositionsLabMain.SetDistributeMarginRight((float)marginRightValue);

            if (alignLeftButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SetDistributeGridAlignment(PositionsLabMain.GridAlignment.AlignLeft);
            }

            if (alignCenterButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SetDistributeGridAlignment(PositionsLabMain.GridAlignment.AlignCenter);
            }

            if (alignRightButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SetDistributeGridAlignment(PositionsLabMain.GridAlignment.AlignRight);
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
