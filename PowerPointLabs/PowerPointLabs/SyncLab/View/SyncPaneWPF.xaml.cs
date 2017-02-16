using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace PowerPointLabs.SyncLab.View
{
    /// <summary>
    /// Interaction logic for SyncPaneWPF.xaml
    /// </summary>
    public partial class SyncPaneWPF : UserControl
    {
#pragma warning disable 0618

        public SyncPaneWPF()
        {
            InitializeComponent();
            copyImage.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    Properties.Resources.LineColor_icon.GetHbitmap(),
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());
            pasteImage.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    Properties.Resources.FillColor_icon.GetHbitmap(),
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());
        }

        #region GUI Handles
        private void CopyImage_MouseLeftButtonUp(object sender, RoutedEventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange.Count != 1)
            {
                MessageBox.Show("Please select one item to copy.");
                return;
            }
            SyncFormatDialog dialog = new SyncFormatDialog();
            bool? result = dialog.ShowDialog();
            MessageBox.Show(result.ToString());
            if (!result.HasValue || !(bool)result)
            {
                return;
            }
            var shape = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
            SyncFormatListItem item = new SyncFormatListItem();
            item.Image = new System.Drawing.Bitmap(Utils.Graphics.ShapeToImage(shape));
            formatListBox.Items.Insert(0, item);
        }

        private void CopyButton_Click(object sender, RoutedEventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange.Count != 1)
            {
                MessageBox.Show("Please select one item to copy.");
                return;
            }
            SyncFormatDialog dialog = new SyncFormatDialog();
            bool? result = dialog.ShowDialog();
            if (!result.HasValue || !(bool)result)
            {
                return;
            }
            var shape = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
            SyncFormatListItem item = new SyncFormatListItem();
            //System.Drawing.Bitmap b = new System.Drawing.Bitmap(30, 30);
            //System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(b);
            //g.DrawImage(Utils.Graphics.ShapeToImage(shape), 0, 0, 30, 30);
            item.Image = new System.Drawing.Bitmap(Utils.Graphics.ShapeToImage(shape)); //b;
            formatListBox.Items.Insert(0, item);
        }

        private void PasteButton_Click(object sender, RoutedEventArgs e)
        {

        }
        #endregion
    }
}
