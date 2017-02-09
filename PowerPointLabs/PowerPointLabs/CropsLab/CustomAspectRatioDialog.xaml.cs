using System.Windows;

namespace PowerPointLabs.CropsLab
{
    /// <summary>
    /// Interaction logic for CustomAspectRatioDialog.xaml
    /// </summary>
    public partial class CustomAspectRatioDialog
    {
#pragma warning disable 0618

        public CustomAspectRatioDialog()
        {
            InitializeComponent();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            string widthText = textBoxWidthInput.Text;
            string heightText = textBoxHeightInput.Text;

            Globals.ThisAddIn.Ribbon.CropToAspectRatioInput(widthText, heightText);
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
