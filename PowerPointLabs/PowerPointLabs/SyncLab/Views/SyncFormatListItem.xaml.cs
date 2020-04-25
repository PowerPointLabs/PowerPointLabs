using System;
using System.Drawing;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media.Imaging;

using PowerPointLabs.Utils;

namespace PowerPointLabs.SyncLab.Views
{
    /// <summary>
    /// Interaction logic for SyncFormatListItem.xaml
    /// </summary>
    public partial class SyncFormatListItem : UserControl
    {
        Bitmap image;

        public SyncFormatListItem()
        {
            InitializeComponent();
            UpdateImage();
        }

        public bool? IsChecked
        {
            get => checkBox.IsChecked;
            set => checkBox.IsChecked = value;
        }

        public Bitmap Image
        {
            get => image;
            set
            {
                image = value;
                UpdateImage();
            }
        }

        private void UpdateImage()
        {
            // if image isn't set, hide the image border.
            if (image == null)
            {
                imageBorder.Visibility = Visibility.Collapsed;
                return;
            }

            imageBorder.Visibility = Visibility.Visible;
            BitmapSource source = GraphicsUtil.BitmapToImageSource(image);
            imageBox.Source = source;
        }

        public string Text
        {
            get => label.Content.ToString();
            set => label.Content = value;
        }
    }
}
