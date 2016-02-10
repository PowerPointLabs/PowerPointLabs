using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MahApps.Metro.Controls;

namespace PowerPointLabs.PositionsLab
{
    /// <summary>
    /// Interaction logic for PositionsPaneWPF.xaml
    /// </summary>
    public partial class PositionsPaneWPF : UserControl
    {

        public PositionsPaneWPF()
        {
            InitializeComponent();
        }

        #region Align
        private void AlignLeftButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.AlignLeft();
        }

        private void AlignRightButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.AlignRight();
        }

        private void AlignTopButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.AlignTop();
        }

        private void AlignBottomButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.AlignBottom();
        }

        private void AlignMiddleButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.AlignMiddle();
        }

        private void AlignCenterButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.AlignCenter();
        }
        #endregion

        #region Adjoin
        private void AdjoinHorizontalButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.AdjoinHorizontal();
        }

        private void AdjoinVerticalButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.AdjoinVertical();
        }
        #endregion

        #region Distribute
        private void DistributeHorizontalButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.DistributeHorizontal();
        }

        private void DistributeVerticalButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.DistributeVertical();
        }
        #endregion

        #region Snap
        private void SnapHorizontalButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.SnapHorizontal();
        }

        private void SnapVerticalButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.SnapVertical();
        }
        #endregion

        #region Swap
        private void SwapPositionsButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.Swap();
        }
        #endregion

    }
}
