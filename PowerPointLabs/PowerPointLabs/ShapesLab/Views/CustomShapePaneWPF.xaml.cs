using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Media.Imaging;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Models;
using PowerPointLabs.ShortcutsLab;
using PowerPointLabs.SyncLab.ObjectFormats;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;
using MessageBox = System.Windows.Forms.MessageBox;
using Point = System.Windows.Point;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;

namespace PowerPointLabs.ShapesLab.Views
{
    /// <summary>
    /// Interaction logic for CustomShapePaneWPF.xaml
    /// </summary>
    public partial class CustomShapePaneWPF : System.Windows.Controls.UserControl
    {
        # region Properties
        public List<string> Categories { get; private set; }

        public string CurrentCategory { get; set; }

        public string CurrentShapeFullName
        {
            get { return CurrentShapeFolderPath + @"\" + CurrentShapeNameWithoutExtension + ".png"; }
        }

        public string CurrentShapeNameWithoutExtension
        {
            get
            {
                if (_selectedThumbnail == null ||
                    _selectedThumbnail.Count == 0)
                {
                    return null;
                }

                return _selectedThumbnail[0].NameLabel;
            }
        }

        public List<string> Shapes
        {
            get
            {
                List<string> shapeList = new List<string>();

                if (wrapPanel.Children.Count == 0)
                {
                    return shapeList;
                }

                //TODO
                //shapeList.AddRange(from LabeledThumbnail labelThumbnail in wrapPanel.Controls
                //                   select labelThumbnail.NameLabel);

                return shapeList;
            }
        }

        public string ShapeRootFolderPath { get; private set; }

        public string CurrentShapeFolderPath
        {
            get { return ShapeRootFolderPath + @"\" + CurrentCategory; }
        }
        
        # endregion

        public CustomShapePaneWPF()
        {
            InitializeComponent();

            copyImage.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    Properties.Resources.AddToCustomShapes.GetHbitmap(),
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());
        }

        public void SyncPaneWPF_Loaded(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Tools.CustomTaskPane syncLabPane = this.GetAddIn().GetActivePane(typeof(CustomShapePane_));
            if (syncLabPane == null || !(syncLabPane.Control is CustomShapePane_))
            {
                MessageBox.Show(TextCollection.SyncLabText.ErrorSyncPaneNotOpened);
                return;
            }
            CustomShapePane_ syncLab = syncLabPane.Control as CustomShapePane_;

            UpdateCopyButtonEnabledStatus(this.GetCurrentSelection());

            syncLab.HandleDestroyed += SyncPane_Closing;
        }

        public void SyncPane_Closing(Object sender, EventArgs e)
        {
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
        
        public string GetFormatText(int index)
        {
            return (formatListBox.Items[index] as CustomShapePaneItem).Text;
        }

        public void SetFormatText(int index, string text)
        {
            (formatListBox.Items[index] as CustomShapePaneItem).Text = text;
        }
        #endregion

        #region API

        public void AddShapeFromSelection(PowerPointSlide currentSlide, PowerPointPresentation pres,
            Selection selection, ThisAddIn addIn)
        {
            // first of all we check if the shape gallery has been opened correctly
            if (!addIn.ShapePresentation.Opened)
            {
                MessageBox.Show(CommonText.ErrorShapeGalleryInit);
                return;
            }

            // Check this so that it is the same requirements as ConvertToPicture which is used when adding shapes
            if (!ShapeUtil.IsSelectionShapeOrText(selection))
            {
                MessageBox.Show(new Form() { TopMost = true },
                    ShapesLabText.ErrorAddSelectionInvalid, ShapesLabText.ErrorDialogTitle);
                return;
            }

            // Finish checks, will add shape(s) from selection

            ShapeRange selectedShapes = selection.ShapeRange;
            if (selection.HasChildShapeRange)
            {
                selectedShapes = selection.ChildShapeRange;
            }

            // add shape into shape gallery first to reduce flicker
            string shapeName = addIn.ShapePresentation.AddShape(pres, currentSlide, selectedShapes, selectedShapes[1].Name);

            // add the selection into pane and save it as .png locally
            string shapeFullName = Path.Combine(CurrentShapeFolderPath, shapeName + ".png");
            bool success = ConvertToPicture.ConvertAndSave(selectedShapes, shapeFullName);
            if (!success)
            {
                return;
            }

            // sync the shape among all opening panels
            addIn.SyncShapeAdd(shapeName, shapeFullName, CurrentCategory);

            // finally, add the shape into the panel and waiting for name editing
            AddCustomShape(shapeName, shapeFullName, true);
        }

        public void AddCustomShape(string shapeName, string shapePath, bool immediateEditing)
        {
            //TODO
            //LabeledThumbnail labeledThumbnail = new LabeledThumbnail(shapePath, shapeName) { ContextMenuStrip = shapeContextMenuStrip };
            wrapPanel.Children.Add(new CustomShapePaneItem(shapeName));

            //labeledThumbnail.Click += LabeledThumbnailClick;
            //labeledThumbnail.DoubleClick += LabeledThumbnailDoubleClick;
            //labeledThumbnail.NameEditFinish += NameEditFinishHandler;

            //ReorderThumbnail(labeledThumbnail);
        }

        public void RemoveCustomShape(string shapeName)
        {
            LabeledThumbnail labeledThumbnail = FindLabeledThumbnail(shapeName);

            if (labeledThumbnail == null)
            {
                return;
            }

            // remove shape from task pane
            RemoveThumbnail(labeledThumbnail);
        }

        public void RenameCustomShape(string oldShapeName, string newShapeName)
        {
            LabeledThumbnail labeledThumbnail = FindLabeledThumbnail(oldShapeName);

            if (labeledThumbnail == null)
            {
                return;
            }

            labeledThumbnail.RenameWithoutEdit(newShapeName);

            //ReorderThumbnail(labeledThumbnail);
        }

        public void UpdateOnSelectionChange(Selection selection)
        {
            SelectionChanged(selection);
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
                CustomShapePaneItem item = formatListBox.Items[index] as CustomShapePaneItem;
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

        private void AddFormatToList(Shape shape, string name)
        {
            //TODO
            CustomShapePaneItem item = new CustomShapePaneItem("");
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
                shape.Type == MsoShapeType.msoPlaceholder;
                //ShapeUtil.CanCopyMsoPlaceHolder(shape, SyncFormatUtil.GetTemplateShapes());

            if (shape.Type != MsoShapeType.msoAutoShape &&
                shape.Type != MsoShapeType.msoLine &&
                shape.Type != MsoShapeType.msoPicture &&
                shape.Type != MsoShapeType.msoTextBox &&
                !canSyncPlaceHolder)
            {
                MessageBox.Show(SyncLabText.ErrorCopySelectionInvalid, SyncLabText.ErrorDialogTitle);
                return;
            }
        }
        #endregion

        #region Shape Saving

        // Saves shape into another powerpoint file
        // Returns a key to find the shape by,
        // or null if the shape cannot be copied
        private string CopyShape(Shape shape)
        {
            //return shapeStorage.CopyShape(shape, GetFormatsToApply(nodes));
            return "";
        }
        #endregion

    }
}