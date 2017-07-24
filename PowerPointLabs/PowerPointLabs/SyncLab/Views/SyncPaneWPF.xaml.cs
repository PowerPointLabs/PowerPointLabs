using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

using Microsoft.Office.Interop.PowerPoint;

using static PowerPointLabs.ActionFramework.Common.Extension.ContentControlExtensions;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.SyncLab.Views
{
    /// <summary>
    /// Interaction logic for SyncPaneWPF.xaml
    /// </summary>
    public partial class SyncPaneWPF : UserControl
    {
        public SyncFormatDialog Dialog { get; set; }
        private readonly SyncLabShapeStorage shapeStorage;

        public SyncPaneWPF()
        {
            InitializeComponent();
            shapeStorage = SyncLabShapeStorage.Instance;

            copyImage.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    Properties.Resources.SyncLabCopyButton.GetHbitmap(),
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());
        }

        public void SyncPaneWPF_Loaded(object sender, RoutedEventArgs e)
        {
            var syncLabPane = this.GetAddIn().GetActivePane(typeof(SyncPane));
            if (syncLabPane == null || !(syncLabPane.Control is SyncPane))
            {
                MessageBox.Show("Error: SyncPane not opened.");
                return;
            }
            var syncLab = syncLabPane.Control as SyncPane;

            syncLab.HandleDestroyed += SyncPane_Closing;
        }

        public void SyncPane_Closing(Object sender, EventArgs e)
        {
            if (this.GetAddIn().Application.Presentations.Count == 2)
            {
                shapeStorage.Close();
            }

            if (Dialog != null)
            {
                Dialog.Close();
            }
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

        public void ApplyFormats(FormatTreeNode[] nodes, Shape formatShape)
        {
            var selection = this.GetCurrentSelection();
            if ((selection.Type != PpSelectionType.ppSelectionShapes &&
                selection.Type != PpSelectionType.ppSelectionText) ||
                selection.ShapeRange.Count == 0)
            {
                MessageBox.Show(SyncLabText.ErrorPasteSelectionInvalid, SyncLabText.ErrorDialogTitle);
                return;
            }

            var shapes = selection.ShapeRange;
            if (selection.HasChildShapeRange)
            {
                shapes = selection.ChildShapeRange;
            }

            ApplyFormats(nodes, formatShape, shapes);
        }

        private void AddFormatToList(Shape shape, string name, FormatTreeNode[] formats)
        {
            string shapeKey = CopyShape(shape);
            if (shapeKey == null)
            {
                MessageBox.Show(SyncLabText.ErrorCopy);
                return;
            }
            SyncFormatPaneItem item = new SyncFormatPaneItem(this, shapeKey, shapeStorage, formats);
            item.Text = name;
            item.Image = new System.Drawing.Bitmap(GraphicsUtil.ShapeToBitmap(shape));
            formatListBox.Items.Insert(0, item);
            formatListBox.SelectedIndex = 0;
        }

        private void ApplyFormats(FormatTreeNode[] nodes, Shape formatShape, ShapeRange newShapes)
        {
            foreach (Shape newShape in newShapes)
            {
                ApplyFormats(nodes, formatShape, newShape);
            }
        }

        private void ApplyFormats(FormatTreeNode[] nodes, Shape formatShape, Shape newShape)
        {
            foreach (FormatTreeNode node in nodes)
            {
                ApplyFormats(node, formatShape, newShape);
            }
        }

        private void ApplyFormats(FormatTreeNode node, Shape formatShape, Shape newShape)
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
            var selection = this.GetCurrentSelection();
            if ((selection.Type != PpSelectionType.ppSelectionShapes &&
                selection.Type != PpSelectionType.ppSelectionText) ||
                selection.ShapeRange.Count != 1)
            {
                MessageBox.Show(SyncLabText.ErrorCopySelectionInvalid, SyncLabText.ErrorDialogTitle);
                return;
            }

            var shape = selection.ShapeRange[1];
            if (selection.HasChildShapeRange)
            {
                if (selection.ChildShapeRange.Count != 1)
                {
                    MessageBox.Show(SyncLabText.ErrorCopySelectionInvalid, SyncLabText.ErrorDialogTitle);
                    return;
                }
                shape = selection.ChildShapeRange[1];
            }

            if (shape.Type != Microsoft.Office.Core.MsoShapeType.msoAutoShape &&
                shape.Type != Microsoft.Office.Core.MsoShapeType.msoLine &&
                shape.Type != Microsoft.Office.Core.MsoShapeType.msoTextBox)
            {
                MessageBox.Show(SyncLabText.ErrorCopySelectionInvalid, SyncLabText.ErrorDialogTitle);
                return;
            }
            Dialog = new SyncFormatDialog(shape);
            Dialog.ObjectName = shape.Name;
            bool? result = Dialog.ShowDialog();
            if (!result.HasValue || !(bool)result)
            {
                return;
            }
            AddFormatToList(shape, Dialog.ObjectName, Dialog.Formats);
            Dialog = null;
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
        #endregion

    }
}