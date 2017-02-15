using System.Windows;

namespace PowerPointLabs.CropLab
{
    /// <summary>
    /// Interaction logic for CustomAspectRatioDialog.xaml
    /// </summary>
    public partial class CustomAspectRatioDialog
    {
        public delegate void UpdateSettingsDelegate(string aspectRatioRawString);
        public UpdateSettingsDelegate SettingsHandler;

        public CustomAspectRatioDialog()
        {
            InitializeComponent();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            SettingsHandler(textBoxWidthInput.Text + ":" + textBoxHeightInput.Text);
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
