using System;
using System.Drawing;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media.Imaging;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.SyncLab.View
{
    /// <summary>
    /// Interaction logic for SyncFormatPaneItem.xaml
    /// </summary>
    public partial class SyncFormatPaneItem : UserControl
    {

        
        Bitmap image;

        SyncPaneWPF parent;
        Shape shape = null;
        FormatTreeNode[] formats = null;

        public SyncFormatPaneItem(SyncPaneWPF parent, Shape shape, FormatTreeNode[] formats)
        {
            InitializeComponent();
            this.parent = parent;
            this.shape = shape;
            this.formats = formats;
        }

        public Shape FormatShape
        {
            get
            {
                return shape;
            }
        }

        public FormatTreeNode[] Formats
        {
            get
            {
                return formats;
            }
            set
            {
                formats = value;
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
                BitmapSource source = Imaging.CreateBitmapSourceFromHBitmap(
                                        image.GetHbitmap(),
                                        IntPtr.Zero,
                                        Int32Rect.Empty,
                                        BitmapSizeOptions.FromEmptyOptions());
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

        private void PasteButton_Click(object sender, RoutedEventArgs e)
        {
            parent.ApplyFormats(formats, shape);
        }

        private void EditButton_Click(object sender, RoutedEventArgs e)
        {
            SyncFormatDialog dialog = new SyncFormatDialog(shape, Text, formats);
            dialog.ObjectName = this.Text;
            bool? result = dialog.ShowDialog();
            if (!result.HasValue || !(bool)result)
            {
                return;
            }
            this.formats = dialog.Formats;
            this.Text = dialog.ObjectName;
        }

        private void OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            parent.ApplyFormats(formats, shape);
        }
    }
}
