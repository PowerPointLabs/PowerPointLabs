using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Utils;

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
            ClearStorageTemplate();
            copyImage.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    Properties.Resources.LineColor_icon.GetHbitmap(),
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
            SyncFormatPaneItem item = new SyncFormatPaneItem(this, CopyShape(shape), formats);
            item.Text = name;
            item.Image = new System.Drawing.Bitmap(Utils.Graphics.ShapeToImage(shape));
            formatListBox.Items.Insert(0, item);
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

        private Shape CopyShape(Shape shape)
        {
            Design design = Graphics.GetDesign(TextCollection.SyncLabStorageTemplateName);
            if (design == null)
            {
                design = Graphics.CreateDesign(TextCollection.SyncLabStorageTemplateName);
            }
            shape.Copy();
            ShapeRange newShapeRange = design.TitleMaster.Shapes.Paste();
            return newShapeRange[1];
        }

        private void ClearStorageTemplate()
        {
            Design design = Graphics.GetDesign(TextCollection.SyncLabStorageTemplateName);
            if (design != null)
            {
                design.Delete();
            }
        }
        #endregion

    }
}