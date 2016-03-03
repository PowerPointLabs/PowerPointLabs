using MahApps.Metro.Controls;
using System.Windows;
using System.Windows.Controls;

namespace PowerPointLabs.PositionsLab
{
    /// <summary>
    /// Interaction logic for DistributeSettingsDialog.xaml
    /// </summary>
    public partial class DistributeSettingsDialog : MetroWindow
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

        private void DistributeToShapeButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.DistributeUseSlideAsReference)
            {
                distributeToShapeButton.IsChecked = false;
            }
        }

        private void DistributeToSlideButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.DistributeUseSlideAsReference)
            {
                distributeToSlideButton.IsChecked = true;
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

            if (distributeToShapeButton.IsChecked == true)
            {
                PositionsLabMain.DistributeReferToShape();
            }
            else
            {
                PositionsLabMain.DistributeReferToSlide();
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

        private void DistributeToShapeButton_Click(object sender, RoutedEventArgs e)
        {
            distributeToSlideButton.IsChecked = false;
        }

        private void DistributeToSlideButton_Click(object sender, RoutedEventArgs e)
        {
            distributeToShapeButton.IsChecked = false;
        }
        #endregion

        private void DistributeSettingsDialong_Closed(object sender, System.EventArgs e)
        {
            IsOpen = false;
        } 
    }
}
