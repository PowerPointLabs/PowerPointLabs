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
            get
            {
                return checkBox.IsChecked;
            }
            set
            {
                checkBox.IsChecked = value;
            }
        }

        public Bitmap Image
        {
            get
            {
                return image;
            }
            set
            {
                image = value;
                UpdateImage();
            }
        }

        private void UpdateImage()
        {
            // if image isn't set, fill the control with the label
            if (image == null)
            {
                imageBox.Visibility = Visibility.Hidden;
                label.Margin = new Thickness(30, label.Margin.Top,
                            label.Margin.Right, label.Margin.Bottom);
                return;
            }
            else
            {
                BitmapSource source = CommonUtil.CreateBitmapSource(image);
                imageBox.Source = source;
                imageBox.Visibility = Visibility.Visible;
                label.Margin = new Thickness(65, label.Margin.Top,
                            label.Margin.Right, label.Margin.Bottom);
            }
        }

        public String Text
        {
            get
            {
                return label.Content.ToString();
            }
            set
            {
                label.Content = value;
            }
        }
    }
}
