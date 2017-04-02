using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;

namespace PowerPointLabs.SyncLab.View
{
    /// <summary>
    /// Interaction logic for SyncPaneWPF.xaml
    /// </summary>
    public partial class SyncPaneWPF : UserControl
    {
#pragma warning disable 0618

        private readonly SyncLabShapeStorage shapeStorage;

        public SyncPaneWPF()
        {
            InitializeComponent();
            shapeStorage = new SyncLabShapeStorage();
            copyImage.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    Properties.Resources.SyncLabCopyButton.GetHbitmap(),
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());
        }

        #region GUI API
        public int FormatCount
        {
            get
            {
                return formatListBox.Items.Count;
            }
        }

        public FormatTreeNode[] GetFormats(int index)
        {
            return (formatListBox.Items[index] as SyncFormatPaneItem).Formats;
        }

        public string GetFormatText(int index)
        {
            return (formatListBox.Items[index] as SyncFormatPaneItem).Text;
        }

        public void SetFormatText(int index, string text)
        {
            (formatListBox.Items[index] as SyncFormatPaneItem).Text = text;
        }
        #endregion

        #region Sync API
        public void AddFormatToList(Shape shape, string name, FormatTreeNode[] formats)
        {
            string shapeKey = CopyShape(shape);
            if (shapeKey == null)
            {
                MessageBox.Show(TextCollection.SyncLabCopyError);
                return;
            }
            SyncFormatPaneItem item = new SyncFormatPaneItem(this, shapeKey, shapeStorage, formats);
            item.Text = name;
            item.Image = new System.Drawing.Bitmap(Utils.Graphics.ShapeToImage(shape));
            formatListBox.Items.Insert(0, item);
            formatListBox.SelectedIndex = 0;
        }

        public void ApplyFormats(FormatTreeNode[] nodes, Shape formatShape)
        {
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type != PpSelectionType.ppSelectionShapes ||
                selection.ShapeRange.Count == 0)
            {
                MessageBox.Show(TextCollection.SyncLabPasteSelectError);
                return;
            }
            ApplyFormats(nodes, formatShape, selection.ShapeRange);
        }

        public void ApplyFormats(FormatTreeNode[] nodes, Shape formatShape, ShapeRange newShapes)
        {
            foreach (Shape newShape in newShapes)
            {
                ApplyFormats(nodes, formatShape, newShape);
            }
        }

        public void ApplyFormats(FormatTreeNode[] nodes, Shape formatShape, Shape newShape)
        {
            foreach (FormatTreeNode node in nodes)
            {
                ApplyFormats(node, formatShape, newShape);
            }
        }

        public void ApplyFormats(FormatTreeNode node, Shape formatShape, Shape newShape)
        {
            if (node.Format != null)
            {
                if (!node.IsChecked.HasValue || !node.IsChecked.Value)
                {
                    return;
                }
                node.Format.SyncFormat(formatShape, newShape);
            }
            else
            {
                ApplyFormats(node.ChildrenNodes, formatShape, newShape);
            }
        }

        public void RemoveFormatItem(Object format)
        {
            int index = 0;
            while (index < formatListBox.Items.Count)
            {
                if (formatListBox.Items[index] == format)
                {
                    formatListBox.Items.RemoveAt(index);
                }
                else
                {
                    index++;
                }
            }
        }

        public void ClearInvalidFormats()
        {
            int index = 0;
            while (index < formatListBox.Items.Count)
            {
                SyncFormatPaneItem item = formatListBox.Items[index] as SyncFormatPaneItem;
                if (item.FormatShapeExists)
                {
                    index++;
                }
                else
                {
                    formatListBox.Items.RemoveAt(index);
                }
            }
        }
        #endregion

        #region GUI Handles
        private void CopyButton_Click(object sender, RoutedEventArgs e)
        {
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type != PpSelectionType.ppSelectionShapes ||
                selection.ShapeRange.Count != 1)
            {
                MessageBox.Show(TextCollection.SyncLabCopySelectError);
                return;
            }
            var shape = selection.ShapeRange[1];
            SyncFormatDialog dialog = new SyncFormatDialog(shape);
            dialog.ObjectName = shape.Name;
            bool? result = dialog.ShowDialog();
            if (!result.HasValue || !(bool)result)
            {
                return;
            }
            AddFormatToList(shape, dialog.ObjectName, dialog.Formats);
        }
        #endregion

        #region Shape Saving

        // Saves shape into another powerpoint file
        // Returns a key to find the shape by,
        // or null if the shape cannot be copied
        private string CopyShape(Shape shape)
        {
            return shapeStorage.CopyShape(shape);
        }

        private void InitializeStoragePresentation()
        {
            PowerPointPresentation presentation = GetStoragePresentation();
            while (presentation.SlideCount > 0)
            {
                presentation.GetSlide(1).Delete();
            }
            presentation.AddSlide();
            presentation.Slides[0].DeleteAllShapes();
        }

        private PowerPointPresentation storagePresentation = null;

        private PowerPointPresentation GetStoragePresentation()
        {
            if (storagePresentation != null)
            {
                return storagePresentation;
            }
            var tempPath = Globals.ThisAddIn.PrepareTempFolder(
                PowerPointPresentation.Current.Presentation);
            var tempName = string.Format(TextCollection.SyncLabStorageTemplateName,
                                         DateTime.Now.ToString("yyyyMMddHHmmssffff"));
            PowerPointPresentation presentation = new PowerPointPresentation(tempPath, tempName);
            storagePresentation = presentation.OpenInBackground() ? presentation : null;
            return storagePresentation;
        }
        #endregion

    }
}