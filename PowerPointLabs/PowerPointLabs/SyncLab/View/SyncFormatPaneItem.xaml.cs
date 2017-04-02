using System;
using System.Drawing;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media.Imaging;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;

namespace PowerPointLabs.SyncLab.View
{
    /// <summary>
    /// Interaction logic for SyncFormatPaneItem.xaml
    /// </summary>
    public partial class SyncFormatPaneItem : UserControl
    {

        private Bitmap image;

        private SyncPaneWPF parent;
        private string shapeKey = null;
        private SyncLabShapeStorage shapeStorage;
        private FormatTreeNode[] formats = null;

        public SyncFormatPaneItem(SyncPaneWPF parent, string shapeKey,
                    SyncLabShapeStorage shapeStorage, FormatTreeNode[] formats)
        {
            InitializeComponent();
            this.parent = parent;
            this.shapeKey = shapeKey;
            this.shapeStorage = shapeStorage;
            this.formats = formats;
            editImage.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                Properties.Resources.SyncLabEditButton.GetHbitmap(),
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
            pasteImage.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                Properties.Resources.SyncLabPasteButton.GetHbitmap(),
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions()); 
            deleteImage.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                Properties.Resources.SyncLabDeleteButton.GetHbitmap(),
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
        }

        public string FormatShapeKey
        {
            get
            {
                return shapeKey;
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

        public bool FormatShapeExists
        {
            get
            {
                return shapeStorage.GetShape(shapeKey) != null;
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
            ApplyFormatToSelected();
        }

        private void EditButton_Click(object sender, RoutedEventArgs e)
        {
            Shape shape = shapeStorage.GetShape(shapeKey);
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

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            parent.RemoveFormatItem(this);
        }

        private void OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ApplyFormatToSelected();
        }

        private void ApplyFormatToSelected()
        {
            Shape formatShape = shapeStorage.GetShape(shapeKey);
            if (formatShape == null)
            {
                MessageBox.Show(TextCollection.SyncLabShapeDeletedError);
                parent.ClearInvalidFormats();
            }
            this.StartNewUndoEntry();
            parent.ApplyFormats(formats, formatShape);
        }
    }
}
