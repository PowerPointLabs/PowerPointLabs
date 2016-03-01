using MahApps.Metro.Controls;
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

namespace PowerPointLabs.PositionsLab
{
    /// <summary>
    /// Interaction logic for DistributeSettingsDialog.xaml
    /// </summary>
    public partial class DistributeSettingsDialog : MetroWindow
    {
        public DistributeSettingsDialog()
        {
            InitializeComponent();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (distributeToShapeButton.IsChecked == true)
            {
                PositionsLabMain.DistributeReferToShape();
            }
            else
            {
                PositionsLabMain.DistributeReferToSlide();
            }
            this.Hide();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
        }

        private void DistributeToShapeButton_Click(object sender, RoutedEventArgs e)
        {
            distributeToSlideButton.IsChecked = false;
        }

        private void DistributeToSlideButton_Click(object sender, RoutedEventArgs e)
        {
            distributeToShapeButton.IsChecked = false;
        }
    }
}
