using MahApps.Metro.Controls;
using System.Windows;

namespace PowerPointLabs.PositionsLab
{
    /// <summary>
    /// Interaction logic for AlignSettingsDialog.xaml
    /// </summary>
    public partial class AlignSettingsDialog : MetroWindow
    {
        //Flag to trigger
        public bool IsOpen { get; set; }

        public AlignSettingsDialog()
        {
            IsOpen = true;
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

            IsOpen = false;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            IsOpen = false;
            Close();
        }

        private void AlignToShapeButton_Click(object sender, RoutedEventArgs e)
        {
            alignToSlideButton.IsChecked = false;
        }

        private void AlignToSlideButton_Click(object sender, RoutedEventArgs e)
        {
            alignToShapeButton.IsChecked = false;
        }

        private void AlignSettingsDialong_Closed(object sender, System.EventArgs e)
        {
            IsOpen = false;
        }
    }
}
