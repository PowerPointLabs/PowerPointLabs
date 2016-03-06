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
        private void SwapByCyclicOrderButton_Load(object sender, RoutedEventArgs e)
        {
            swapByCyclicOrderButton.IsChecked = !PositionsLabMain.SwapByClickOrder;
        }

        private void SwapByClickOrderButton_Load(object sender, RoutedEventArgs e)
        {
            swapByClickOrderButton.IsChecked = PositionsLabMain.SwapByClickOrder;
        }

        private void CenterAsReferenceButton_Load(object sender, RoutedEventArgs e)
        {
            centerAsReferenceButton.IsChecked = PositionsLabMain.SwapReferencePoint ==
                                                PositionsLabMain.SwapReference.Center;
        }

        private void LeftAsReferenceButton_Load(object sender, RoutedEventArgs e)
        {
            leftAsReferenceButton.IsChecked = PositionsLabMain.SwapReferencePoint ==
                                                PositionsLabMain.SwapReference.Left;
        }

        private void TopAsReferenceButton_Load(object sender, RoutedEventArgs e)
        {
            topAsReferenceButton.IsChecked = PositionsLabMain.SwapReferencePoint ==
                                                PositionsLabMain.SwapReference.Top;
        }

        private void RightAsReferenceButton_Load(object sender, RoutedEventArgs e)
        {
            rightAsReferenceButton.IsChecked = PositionsLabMain.SwapReferencePoint ==
                                                PositionsLabMain.SwapReference.Right;
        }

        private void BottomAsReferenceButton_Load(object sender, RoutedEventArgs e)
        {
            bottomAsReferenceButton.IsChecked = PositionsLabMain.SwapReferencePoint ==
                                                PositionsLabMain.SwapReference.Bottom;
        }
        #endregion

        #region Button Actions
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {

            PositionsLabMain.SwapByClickOrder = swapByCyclicOrderButton.IsChecked.GetValueOrDefault();
            
            if (centerAsReferenceButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SwapReferencePoint = PositionsLabMain.SwapReference.Center;
            }

            if (leftAsReferenceButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SwapReferencePoint = PositionsLabMain.SwapReference.Left;
            }

            if (topAsReferenceButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SwapReferencePoint = PositionsLabMain.SwapReference.Top;
            }

            if (rightAsReferenceButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SwapReferencePoint = PositionsLabMain.SwapReference.Right;
            }

            if (bottomAsReferenceButton.IsChecked.GetValueOrDefault())
            {
                PositionsLabMain.SwapReferencePoint = PositionsLabMain.SwapReference.Bottom;
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
