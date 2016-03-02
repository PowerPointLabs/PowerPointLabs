using MahApps.Metro.Controls;
using System.Windows;

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
