using System;
using System.Windows;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.CropLab
{
    /// <summary>
    /// Interaction logic for CustomAspectRatioDialog.xaml
    /// </summary>
    public partial class CustomAspectRatioDialog
    {
        public delegate void UpdateSettingsDelegate(string aspectRatioRawString);
        public UpdateSettingsDelegate SettingsHandler;

        public CustomAspectRatioDialog(Shape refShape = null)
        {
            InitializeComponent();

            if (refShape != null)
            {
                textBoxWidthInput.Text = Math.Round(refShape.Width / refShape.Height, 4).ToString();
                textBoxHeightInput.Text = "1";
            }
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
