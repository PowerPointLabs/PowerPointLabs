using System;
using System.Windows;

namespace PowerPointLabs.PositionsLab
{
    /// <summary>
    /// Interaction logic for ReorderSettingsDialog.xaml
    /// </summary>
    public partial class ReorderSettingsDialog
    {
        //Flag to trigger
        public bool IsOpen { get; set; }

        public ReorderSettingsDialog()
        {
            IsOpen = true;
            InitializeComponent();
        }

        #region On-Load Settings
        private void SwapByLeftToRightButton_Load(object sender, RoutedEventArgs e)
        {
            swapByLeftToRightButton.IsChecked = !PositionsLabMain.SwapByClickOrder;
        }

        private void SwapByClickOrderButton_Load(object sender, RoutedEventArgs e)
        {
            swapByClickOrderButton.IsChecked = PositionsLabMain.SwapByClickOrder;
        }

        private void TopLeftAsReferenceButton_Load(object sender, RoutedEventArgs e)
        {
            topLeftAsReferenceButton.IsChecked = PositionsLabMain.SwapReferencePoint ==
                                                PositionsLabMain.SwapReference.TopLeft;
        }

        private void TopCenterAsReferenceButton_Load(object sender, RoutedEventArgs e)
        {
            topCenterAsReferenceButton.IsChecked = PositionsLabMain.SwapReferencePoint ==
                                                PositionsLabMain.SwapReference.TopCenter;
        }

        private void TopRightAsReferenceButton_Load(object sender, RoutedEventArgs e)
        {
            topRightAsReferenceButton.IsChecked = PositionsLabMain.SwapReferencePoint ==
                                                PositionsLabMain.SwapReference.TopRight;
        }

        private void MiddleLeftAsReferenceButton_Load(object sender, RoutedEventArgs e)
        {
            middleLeftAsReferenceButton.IsChecked = PositionsLabMain.SwapReferencePoint ==
                                                PositionsLabMain.SwapReference.MiddleLeft;
        }

        private void MiddleCenterAsReferenceButton_Load(object sender, RoutedEventArgs e)
        {
            middleCenterAsReferenceButton.IsChecked = PositionsLabMain.SwapReferencePoint ==
                                                PositionsLabMain.SwapReference.MiddleCenter;
        }

        private void MiddleRightAsReferenceButton_Load(object sender, RoutedEventArgs e)
        {
            middleRightAsReferenceButton.IsChecked = PositionsLabMain.SwapReferencePoint ==
                                                PositionsLabMain.SwapReference.MiddleRight;
        }

        private void BottomLeftAsReferenceButton_Load(object sender, RoutedEventArgs e)
        {
            bottomLeftAsReferenceButton.IsChecked = PositionsLabMain.SwapReferencePoint ==
                                                PositionsLabMain.SwapReference.BottomLeft;
        }

        private void BottomCenterAsReferenceButton_Load(object sender, RoutedEventArgs e)
        {
            bottomCenterAsReferenceButton.IsChecked = PositionsLabMain.SwapReferencePoint ==
                                                PositionsLabMain.SwapReference.BottomCenter;
        }

        private void BottomRightAsReferenceButton_Load(object sender, RoutedEventArgs e)
        {
            bottomRightAsReferenceButton.IsChecked = PositionsLabMain.SwapReferencePoint ==
                                                PositionsLabMain.SwapReference.BottomRight;
        }
        #endregion

        #region Button Actions
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {

            PositionsLabMain.SwapByClickOrder = swapByClickOrderButton.IsChecked.GetValueOrDefault();

            if (topLeftAsReferenceButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SwapReferencePoint = PositionsLabMain.SwapReference.TopLeft;
            }

            if (topCenterAsReferenceButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SwapReferencePoint = PositionsLabMain.SwapReference.TopCenter;
            }

            if (topRightAsReferenceButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SwapReferencePoint = PositionsLabMain.SwapReference.TopRight;
            }

            if (middleLeftAsReferenceButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SwapReferencePoint = PositionsLabMain.SwapReference.MiddleLeft;
            }

            if (middleCenterAsReferenceButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SwapReferencePoint = PositionsLabMain.SwapReference.MiddleCenter;
            }

            if (middleRightAsReferenceButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SwapReferencePoint = PositionsLabMain.SwapReference.MiddleRight;
            }

            if (bottomLeftAsReferenceButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SwapReferencePoint = PositionsLabMain.SwapReference.BottomLeft;
            }

            if (bottomCenterAsReferenceButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SwapReferencePoint = PositionsLabMain.SwapReference.BottomCenter;
            }

            if (bottomRightAsReferenceButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SwapReferencePoint = PositionsLabMain.SwapReference.BottomRight;
            }

            IsOpen = false;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void ReorderSettingsDialong_Closed(object sender, EventArgs e)
        {
            IsOpen = false;
        }        
        #endregion
        

    }
}
