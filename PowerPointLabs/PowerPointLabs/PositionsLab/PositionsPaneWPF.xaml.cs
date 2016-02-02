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

        private void SnapHorizontalButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.SnapHorizontal();
        }

        private void SnapVerticalButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.SnapVertical();
        }

        private void AlignLeftButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.AlignLeft();
        }

        private void AlignRightButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.AlignRight();
        }
    }
}
