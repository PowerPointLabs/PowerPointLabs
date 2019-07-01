using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ColorThemes.Extensions;
using PowerPointLabs.SyncLab.ObjectFormats;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;

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

            copyImage.Source = GraphicsUtil.BitmapToImageSource(Properties.Resources.SyncLabCopyButton);
        }

        public void SyncPaneWPF_Loaded(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Tools.CustomTaskPane syncLabPane = this.GetAddIn().GetActivePane(typeof(SyncPane));
            if (syncLabPane == null || !(syncLabPane.Control is SyncPane))
            {
                MessageBox.Show(TextCollection.SyncLabText.ErrorSyncPaneNotOpened);
                return;
            }
            SyncPane syncLab = syncLabPane.Control as SyncPane;

            UpdateCopyButtonEnabledStatus(this.GetCurrentSelection());

            syncLab.HandleDestroyed += SyncPane_Closing;
        }

        public void SyncPane_Closing(Object sender, EventArgs e)
        {
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

        public void UpdateCopyButtonEnabledStatus(Selection selection)
        {
            if ((selection == null) || (selection.Type == PpSelectionType.ppSelectionNone) 
                || (selection.Type == PpSelectionType.ppSelectionSlides))
            {
                copyButton.IsEnabled = false;
                toolTipTextBox.Text = SyncLabText.DisabledToolTipText;
            }
            else
            {
                copyButton.IsEnabled = true;
                toolTipTextBox.Text = SyncLabText.EnabledToolTipText;
            }
        }

        public bool GetCopyButtonEnabledStatus()
        {
            return copyButton.IsEnabled;
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

        /// <summary>
        /// Applies a set of formats from a source shape to shapes selected by the user 
        /// </summary>
        /// <param name="nodes"></param>
        /// <param name="formatShape">source shape</param>
        public void ApplyFormats(FormatTreeNode[] nodes, Shape formatShape)
        {
            ShapeRange selectedShapes = GetSelectedShapesForFormatting();
            if (selectedShapes == null)
            {
                MessageBox.Show(SyncLabText.ErrorPasteSelectionInvalid, SyncLabText.ErrorDialogTitle);
            }
            else
            {
                Format[] formats = GetFormatsToApply(nodes);
                ShapeUtil.ApplyFormats(formats, formatShape, selectedShapes);
                
            }
        }
        
        private Format[] GetFormatsToApply(FormatTreeNode[] nodes)
        {
            List<Format> list = new List<Format>();
            foreach (FormatTreeNode node in nodes)
            {
                if (node.IsFormatNode && node.IsChecked.HasValue && node.IsChecked.Value)
                {
                    list.Add(node.Format);
                }
                else if (node.ChildrenNodes != null)
                {
                    list.AddRange(GetFormatsToApply(node.ChildrenNodes));
                }
            }

            return list.ToArray();
        }

        /// <summary>
        /// Get shapes selected by user
        /// </summary>
        /// <returns>ShapeRange of selected shapes, or null.
        /// Null is returned over an empty collection as selections may not contain ShapeRanges
        /// </returns>
        private ShapeRange GetSelectedShapesForFormatting()
        {
            Selection selection = this.GetCurrentSelection();
            if ((selection.Type != PpSelectionType.ppSelectionShapes &&
                selection.Type != PpSelectionType.ppSelectionText) ||
                selection.ShapeRange.Count == 0)
            {
                return null;
            }

            ShapeRange shapes = selection.ShapeRange;
            if (selection.HasChildShapeRange)
            {
                shapes = selection.ChildShapeRange;
            }

            return shapes;
        }

        private void AddFormatToList(Shape shape, string name, FormatTreeNode[] formats)
        {
            string shapeKey = CopyShape(shape, formats);
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
        #endregion

        #region GUI Handles
        private void CopyButton_Click(object sender, RoutedEventArgs e)
        {
            Selection selection = this.GetCurrentSelection();
            if ((selection.Type != PpSelectionType.ppSelectionShapes &&
                selection.Type != PpSelectionType.ppSelectionText) ||
                selection.ShapeRange.Count != 1)
            {
                MessageBox.Show(SyncLabText.ErrorCopySelectionInvalid, SyncLabText.ErrorDialogTitle);
                return;
            }

            Shape shape = selection.ShapeRange[1];

            if (shape.Type == MsoShapeType.msoSmartArt) 
            {
                MessageBox.Show(SyncLabText.ErrorSmartArtUnsupported, SyncLabText.ErrorDialogTitle);
                return;
            }
            
            if (selection.HasChildShapeRange)
            {
                if (selection.ChildShapeRange.Count != 1)
                {
                    MessageBox.Show(SyncLabText.ErrorCopySelectionInvalid, SyncLabText.ErrorDialogTitle);
                    return;
                }
                shape = selection.ChildShapeRange[1];
            }

            bool canSyncPlaceHolder =
                shape.Type == MsoShapeType.msoPlaceholder && 
                ShapeUtil.CanCopyMsoPlaceHolder(shape, SyncFormatUtil.GetTemplateShapes());

            if (shape.Type != MsoShapeType.msoAutoShape &&
                shape.Type != MsoShapeType.msoLine &&
                shape.Type != MsoShapeType.msoPicture &&
                shape.Type != MsoShapeType.msoTextBox &&
                !canSyncPlaceHolder)
            {
                MessageBox.Show(SyncLabText.ErrorCopySelectionInvalid, SyncLabText.ErrorDialogTitle);
                return;
            }
            Dialog = new SyncFormatDialog(shape);
            Dialog.ObjectName = shape.Name;
            bool? result = Dialog.ShowThematicDialog();
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
        private string CopyShape(Shape shape, FormatTreeNode[] nodes)
        {
            return shapeStorage.CopyShape(shape, GetFormatsToApply(nodes));
        }
        #endregion

    }
}