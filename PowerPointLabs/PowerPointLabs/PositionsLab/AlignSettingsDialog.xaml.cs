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

        private void AlignToSlideButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.AlignUseSlideAsReference)
            {
                alignToSlideButton.IsChecked = true;
            }
        }

        private void AlignToShapeButton_Load(object sender, RoutedEventArgs e)
        {
            if (PositionsLabMain.AlignUseSlideAsReference)
            {
                alignToShapeButton.IsChecked = false;
            }
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

        private void AlignSettingsDialong_Closed(object sender, System.EventArgs e)
        {
            IsOpen = false;
        }
    }
}
