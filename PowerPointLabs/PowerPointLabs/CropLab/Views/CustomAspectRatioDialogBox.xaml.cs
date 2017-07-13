using System;
using System.Windows;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.CropLab.Views
{
    /// <summary>
    /// Interaction logic for CustomAspectRatioDialogBox.xaml
    /// </summary>
    public partial class CustomAspectRatioDialogBox
    {
        public delegate void UpdateSettingsDelegate(string aspectRatioRawString);
        public UpdateSettingsDelegate SettingsHandler { get; set; }

        public CustomAspectRatioDialogBox(Shape refShape = null)
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
