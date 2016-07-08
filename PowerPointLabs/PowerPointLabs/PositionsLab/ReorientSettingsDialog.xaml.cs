using System.Windows;

namespace PowerPointLabs.PositionsLab
{
    /// <summary>
    /// Interaction logic for ReorientSettingsDialog.xaml
    /// </summary>
    public partial class ReorientSettingsDialog
    {
        //Flag to trigger
        public bool IsOpen { get; set; }

        public ReorientSettingsDialog()
        {
            IsOpen = true;
            InitializeComponent();
        }

        #region On-Load Settings
        private void RadialShapeOrientationFixedButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.RadialShapeOrientation == PositionsLabMain.RadialShapeOrientationObject.Fixed)
            {
                radialShapeOrientationFixedButton.IsChecked = true;
            }
        }

        private void RadialShapeOrientationDynamicButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.RadialShapeOrientation == PositionsLabMain.RadialShapeOrientationObject.Dynamic)
            {
                radialShapeOrientationDynamicButton.IsChecked = true;
            }
        }
        #endregion

        #region Button actions
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            // Checks for radial shape orientation
            if (radialShapeOrientationFixedButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.RadialShapeOrientationToFixed();
            }

            if (radialShapeOrientationDynamicButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.RadialShapeOrientationToDynamic();
            }

            IsOpen = false;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            IsOpen = false;
            Close();
        }

        private void ReorientSettingsDialong_Closed(object sender, System.EventArgs e)
        {
            IsOpen = false;
        }
        #endregion
    }
}
