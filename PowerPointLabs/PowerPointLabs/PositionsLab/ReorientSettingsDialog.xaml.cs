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
        private void ReorientShapeOrientationFixedButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.ReorientShapeOrientation == PositionsLabMain.RadialShapeOrientationObject.Fixed)
            {
                reorientShapeOrientationFixedButton.IsChecked = true;
            }
        }

        private void ReorientShapeOrientationDynamicButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.ReorientShapeOrientation == PositionsLabMain.RadialShapeOrientationObject.Dynamic)
            {
                reorientShapeOrientationDynamicButton.IsChecked = true;
            }
        }
        #endregion

        #region Button actions
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            // Checks for radial shape orientation
            if (reorientShapeOrientationFixedButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.ReorientShapeOrientationToFixed();
            }

            if (reorientShapeOrientationDynamicButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.ReorientShapeOrientationToDynamic();
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
