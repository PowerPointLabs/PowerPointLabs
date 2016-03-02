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
    /// Interaction logic for AlignSettingsDialog.xaml
    /// </summary>
    public partial class AlignSettingsDialog : MetroWindow
    {
        public AlignSettingsDialog()
        {
            InitializeComponent();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (alignToShapeButton.IsChecked == true)
            {
                PositionsLabMain.AlignReferToShape();
            }
            else
            {
                PositionsLabMain.AlignReferToSlide();
            }
            this.Hide();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
        }

        private void AlignToShapeButton_Click(object sender, RoutedEventArgs e)
        {
            alignToSlideButton.IsChecked = false;
        }

        private void AlignToSlideButton_Click(object sender, RoutedEventArgs e)
        {
            alignToShapeButton.IsChecked = false;
        }
    }
}
