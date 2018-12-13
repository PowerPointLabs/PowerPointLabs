using System.Windows;

namespace PowerPointLabs.PositionsLab.Views
{
    /// <summary>
    /// Interaction logic for ReorderSettingsDialog.xaml
    /// </summary>
    public partial class ReorderSettingsDialog
    {
        public delegate void DialogConfirmedDelegate(bool isSwapByClickOrder, PositionsLabSettings.SwapReference swapReferencePoint);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

        public ReorderSettingsDialog()
        {
            InitializeComponent();
        }

        public ReorderSettingsDialog(bool isSwapByClickOrder, PositionsLabSettings.SwapReference swapReferencePoint)
            : this()
        {
            swapByLeftToRightButton.IsChecked = !isSwapByClickOrder;
            swapByClickOrderButton.IsChecked = isSwapByClickOrder;

            switch (swapReferencePoint)
            {
                case PositionsLabSettings.SwapReference.TopLeft:
                    topLeftAsReferenceButton.IsChecked = true;
                    break;
                case PositionsLabSettings.SwapReference.TopCenter:
                    topCenterAsReferenceButton.IsChecked = true;
                    break;
                case PositionsLabSettings.SwapReference.TopRight:
                    topRightAsReferenceButton.IsChecked = true;
                    break;
                case PositionsLabSettings.SwapReference.MiddleLeft:
                    middleLeftAsReferenceButton.IsChecked = true;
                    break;
                case PositionsLabSettings.SwapReference.MiddleCenter:
                    middleCenterAsReferenceButton.IsChecked = true;
                    break;
                case PositionsLabSettings.SwapReference.MiddleRight:
                    middleRightAsReferenceButton.IsChecked = true;
                    break;
                case PositionsLabSettings.SwapReference.BottomLeft:
                    bottomLeftAsReferenceButton.IsChecked = true;
                    break;
                case PositionsLabSettings.SwapReference.BottomCenter:
                    bottomCenterAsReferenceButton.IsChecked = true;
                    break;
                case PositionsLabSettings.SwapReference.BottomRight:
                    bottomRightAsReferenceButton.IsChecked = true;
                    break;
            }
        }
        
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabSettings.SwapReference swapReference;

            if (topLeftAsReferenceButton.IsChecked.GetValueOrDefault())
            {
                swapReference = PositionsLabSettings.SwapReference.TopLeft;
            }
            else if (topCenterAsReferenceButton.IsChecked.GetValueOrDefault())
            {
                swapReference = PositionsLabSettings.SwapReference.TopCenter;
            }
            else if (topRightAsReferenceButton.IsChecked.GetValueOrDefault())
            {
                swapReference = PositionsLabSettings.SwapReference.TopRight;
            }
            else if (middleLeftAsReferenceButton.IsChecked.GetValueOrDefault())
            {
                swapReference = PositionsLabSettings.SwapReference.MiddleLeft;
            }
            else if (middleCenterAsReferenceButton.IsChecked.GetValueOrDefault())
            {
                swapReference = PositionsLabSettings.SwapReference.MiddleCenter;
            }
            else if (middleRightAsReferenceButton.IsChecked.GetValueOrDefault())
            {
                swapReference = PositionsLabSettings.SwapReference.MiddleRight;
            }
            else if (bottomLeftAsReferenceButton.IsChecked.GetValueOrDefault())
            {
                swapReference = PositionsLabSettings.SwapReference.BottomLeft;
            }
            else if (bottomCenterAsReferenceButton.IsChecked.GetValueOrDefault())
            {
                swapReference = PositionsLabSettings.SwapReference.BottomCenter;
            }
            else
            {
                swapReference = PositionsLabSettings.SwapReference.BottomRight;
            }
            
            DialogConfirmedHandler(swapByClickOrderButton.IsChecked.GetValueOrDefault(), swapReference);
            Close();
        }
    }
}
