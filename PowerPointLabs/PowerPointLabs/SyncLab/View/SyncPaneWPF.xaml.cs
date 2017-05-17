using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

using Microsoft.Office.Interop.PowerPoint;

using static PowerPointLabs.ActionFramework.Common.Extension.ContentControlExtensions;
using PowerPointLabs.SyncLab.ObjectFormats;

namespace PowerPointLabs.SyncLab.View
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
            this.Loaded += SyncPaneWPF_Loaded;
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
            shapeStorage.Close();
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

        public SyncFormatDialog ShowDialog(Shape shape)
        {
            return ShowDialog(shape, shape.Name, SyncFormatConstants.FormatCategories);
        }

        public SyncFormatDialog ShowDialog(Shape shape, String formatName, FormatTreeNode[] formats)
        {
            if (Dialog != null)
            {
                Dialog.Close();
            }

            Dialog = new SyncFormatDialog(this, shape, formatName, formats);
            Dialog.Show();
            return Dialog;
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
            if (selection.Type != PpSelectionType.ppSelectionShapes ||
                selection.ShapeRange.Count == 0)
            {
                MessageBox.Show(TextCollection.SyncLabPasteSelectError, TextCollection.SyncLabErrorDialogTitle);
                return;
            }
            ApplyFormats(nodes, formatShape, selection.ShapeRange);
        }

        private void Subscribe(SyncFormatDialog eventDialog)
        {
            eventDialog.OkButtonClick += new SyncFormatDialog.OkButtonEventHandler(AddFormatToList);
        }

        private void AddFormatToList(SyncFormatDialog eventDialog)
        {
            Shape shape = eventDialog.Shape;
            string name = eventDialog.OriginalName;
            FormatTreeNode[] formats = eventDialog.Formats;

            string shapeKey = CopyShape(shape);
            if (shapeKey == null)
            {
                MessageBox.Show(TextCollection.SyncLabCopyError);
                return;
            }
            SyncFormatPaneItem item = new SyncFormatPaneItem(this, shapeKey, shapeStorage, formats);
            item.Text = name;
            item.Image = new System.Drawing.Bitmap(Utils.Graphics.ShapeToBitmap(shape));
            formatListBox.Items.Insert(0, item);
            formatListBox.SelectedIndex = 0;
            eventDialog.OkButtonClick -= new SyncFormatDialog.OkButtonEventHandler(AddFormatToList);
            Dialog = null;
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
            if (selection.Type != PpSelectionType.ppSelectionShapes ||
                selection.ShapeRange.Count != 1)
            {
                MessageBox.Show(TextCollection.SyncLabCopySelectError, TextCollection.SyncLabErrorDialogTitle);
                return;
            }
            var shape = selection.ShapeRange[1];
            Subscribe(ShowDialog(shape));
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