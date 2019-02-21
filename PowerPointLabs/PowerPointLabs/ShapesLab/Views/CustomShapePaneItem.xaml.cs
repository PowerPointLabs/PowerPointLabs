using System;
using System.Drawing;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media.Imaging;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ShapesLab.Views
{
    /// <summary>
    /// Interaction logic for CustomShapePaneItem.xaml
    /// </summary>
    public partial class CustomShapePaneItem : UserControl
    {

        private Bitmap image;

        private string shapeKey = null;

        public CustomShapePaneItem(string shapeKey)
        {
            InitializeComponent();
            this.shapeKey = shapeKey;
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
                return true;
            }
        }

        private void UpdateImage()
        {
            // if image isn't set, fill the control with the label
            if (image == null)
            {
                imageBox.Visibility = Visibility.Hidden;
                col1.Width = new GridLength(0);
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
                col1.Width = new GridLength(60);
            }
        }

        public String Text
        {
            get
            {
                return textBlock.Text;
            }
            set
            {
                textBlock.Text = value;
                textBlock.ToolTip = value;
            }
        }

        private void PasteButton_Click(object sender, RoutedEventArgs e)
        {
            ApplyFormatToSelected();
        }

        private void EditButton_Click(object sender, RoutedEventArgs e)
        {
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            //TODO
            //parent.RemoveFormatItem(this);
        }

        private void OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ApplyFormatToSelected();
        }

        private void ApplyFormatToSelected()
        {
            MessageBox.Show(SyncLabText.ErrorShapeDeleted, SyncLabText.ErrorDialogTitle);
            this.StartNewUndoEntry();
        }
    }
}
