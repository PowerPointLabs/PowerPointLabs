using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Media.Imaging;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Models;
using PowerPointLabs.ShortcutsLab;
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
        private const string ImportFileNameNoExtension = "import";
        private const string ImportFileCopyName = ImportFileNameNoExtension + ".pptx";

        private readonly Comparers.AtomicNumberStringCompare _stringComparer = new Comparers.AtomicNumberStringCompare();

        private BindingSource _categoryBinding;
        private WrapPanel wrapPanel;

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
                return "";
                //TODO
                /*
                if (_selectedThumbnail == null ||
                    _selectedThumbnail.Count == 0)
                {
                    return null;
                }

                return _selectedThumbnail[0].NameLabel;*/
            }
        }

        public string ShapeRootFolderPath { get; private set; }

        public string CurrentShapeFolderPath
        {
            get { return ShapeRootFolderPath + @"\" + CurrentCategory; }
        }

        public bool IsStorageSettingsGiven
        {
            get
            {
                return ShapeRootFolderPath != null && CurrentCategory != null;
            }
        }

        #endregion

        #region Constructors

        public CustomShapePaneWPF()
        {
            InitializeComponent();

            addShapeImage.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    Properties.Resources.AddToCustomShapes.GetHbitmap(),
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());

            InitializeContextMenu();

            //TODO
            /*
            myShapeFlowLayout.MouseEnter += FlowLayoutMouseEnterHandler;
            myShapeFlowLayout.MouseDown += FlowLayoutMouseDownHandler;
            myShapeFlowLayout.MouseUp += FlowLayoutMouseUpHandler;
            myShapeFlowLayout.MouseMove += FlowLayoutMouseMoveHandler;
            */

            //singleShapeDownloadLink.LinkClicked += (s, e) => Process.Start(CommonText.SingleShapeDownloadUrl);

        }

        #endregion

        #region init

        public void SetStorageSettings(string shapeRootFolderPath, string defaultShapeCategoryName)
        {
            ShapeRootFolderPath = shapeRootFolderPath;

            CurrentCategory = defaultShapeCategoryName;
            Categories = new List<string>(Globals.ThisAddIn.ShapePresentation.Categories);
            _categoryBinding = new BindingSource { DataSource = Categories };
            categoryBox.DataContext = _categoryBinding;

            for (int i = 0; i < Categories.Count; i++)
            {
                if (Categories[i] == defaultShapeCategoryName)
                {
                    categoryBox.SelectedIndex = i;
                    break;
                }
            }

            PaneReload();
        }

        public void CustomShapePaneWPF_Loaded(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Tools.CustomTaskPane shapesLabPane = this.GetAddIn().GetActivePane(typeof(CustomShapePane));
            CustomShapePane customShapePane = shapesLabPane?.Control as CustomShapePane;

            if (customShapePane == null)
            {
                MessageBox.Show(TextCollection.SyncLabText.ErrorSyncPaneNotOpened);
                return;
            }

            UpdateAddShapeButtonEnabledStatus(this.GetCurrentSelection());
            customShapePane.HandleDestroyed += CustomShapePane_Closing;
        }

        public void CustomShapePane_Closing(Object sender, EventArgs e)
        {
        }

        #endregion

        #region API

        public void UpdateAddShapeButtonEnabledStatus(Selection selection)
        {
            if ((selection == null) || (selection.Type == PpSelectionType.ppSelectionNone)
                || (selection.Type == PpSelectionType.ppSelectionSlides)
                || !ShapeUtil.IsSelectionShapeOrText(selection))
            {
                DisableAddShapesButton();
            }
            else
            {
                EnableAddShapesButton();
            }
        }

        public bool GetAddShapeButtonEnabledStatus()
        {
            return addShapeButton.IsEnabled;
        }
        
        public string GetShapeName(int index)
        {
            return (shapeList.Items[index] as CustomShapePaneItem).Text;
        }

        public void SetShapeName(int index, string text)
        {
            (shapeList.Items[index] as CustomShapePaneItem).Text = text;
        }

        public void AddShapeFromSelection(Selection selection, ThisAddIn addIn)
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
            PowerPointSlide currentSlide = this.GetCurrentSlide();
            PowerPointPresentation pres = this.GetCurrentPresentation();
            string shapeName = addIn.ShapePresentation.AddShape(pres, currentSlide, selectedShapes, selectedShapes[1].Name);

            // add the selection into pane and save it as .png locally
            string shapePath = Path.Combine(CurrentShapeFolderPath, shapeName + ".png");
            bool success = ConvertToPicture.ConvertAndSave(selectedShapes, shapePath);
            if (!success)
            {
                return;
            }

            // sync the shape among all opening panels
            addIn.SyncShapeAdd(shapeName, shapePath, CurrentCategory);

            // finally, add the shape into the panel and waiting for name editing
            AddCustomShape(shapeName, shapePath, true);
        }

        /// <summary>
        /// Adds a shape lexicographically.
        /// </summary>
        public void AddCustomShape(string shapeName, string shapePath, bool immediateEditing)
        {
            DehighlightSelected();

            //TODO
            //LabeledThumbnail labeledThumbnail = new LabeledThumbnail(shapePath, shapeName) { ContextMenuStrip = shapeContextMenuStrip };
            CustomShapePaneItem shapeItem = new CustomShapePaneItem(shapeName, shapePath);

            //shapeItem.Image = new System.Drawing.Bitmap(GraphicsUtil.ShapeToBitmap(shape));
            int insertionIndex = GetShapeInsertionIndex(shapeName);
            shapeList.Items.Insert(insertionIndex, shapeItem);
            shapeList.SelectedIndex = insertionIndex;

            //TODO
            //labeledThumbnail.Click += LabeledThumbnailClick;
            //labeledThumbnail.DoubleClick += LabeledThumbnailDoubleClick;
            //labeledThumbnail.NameEditFinish += NameEditFinishHandler;
        }

        public void RemoveCustomShape(string shapeName)
        {
            int shapeIndex = GetShapeItemIndex(shapeName);
            shapeList.Items.RemoveAt(shapeIndex);
        }

        public void RenameCustomShape(string oldShapeName, string newShapeName)
        {
            int shapeIndex = GetShapeItemIndex(oldShapeName);

            //TODO
            //shapeItem?.RenameWithoutEdit(newShapeName);
        }

        #endregion

        #region Context Menu

        private void InitializeContextMenu()
        {
            //TODO
            /*
            addToSlideToolStripMenuItem.Text = ShapesLabText.ShapeContextStripAddToSlide;
            editNameToolStripMenuItem.Text = ShapesLabText.ShapeContextStripEditName;
            moveShapeToolStripMenuItem.Text = ShapesLabText.ShapeContextStripMoveShape;
            removeShapeToolStripMenuItem.Text = ShapesLabText.ShapeContextStripRemoveShape;
            copyToToolStripMenuItem.Text = ShapesLabText.ShapeContextStripCopyShape;

            addCategoryToolStripMenuItem.Text = ShapesLabText.CategoryContextStripAddCategory;
            removeCategoryToolStripMenuItem.Text = ShapesLabText.CategoryContextStripRemoveCategory;
            renameCategoryToolStripMenuItem.Text = ShapesLabText.CategoryContextStripRenameCategory;
            setAsDefaultToolStripMenuItem.Text = ShapesLabText.CategoryContextStripSetAsDefaultCategory;
            settingsToolStripMenuItem.Text = ShapesLabText.CategoryContextStripCategorySettings;
            importCategoryToolStripMenuItem.Text = ShapesLabText.CategoryContextStripImportCategory;
            importShapesToolStripMenuItem.Text = ShapesLabText.CategoryContextStripImportShapes;

            foreach (ToolStripMenuItem contextMenu in shapeContextMenuStrip.Items)
            {
                if (contextMenu.Text != ShapesLabText.ShapeContextStripMoveShape)
                {
                    contextMenu.MouseEnter += MoveContextMenuStripLeaveEvent;
                }

                if (contextMenu.Text != ShapesLabText.ShapeContextStripCopyShape)
                {
                    contextMenu.MouseEnter += CopyContextMenuStripLeaveEvent;
                }
            }
            */
        }

        private void ContextMenuStripAddClicked()
        {
            this.StartNewUndoEntry();

            PowerPointSlide currentSlide = this.GetCurrentSlide();
            PowerPointPresentation pres = this.GetCurrentPresentation();

            if (currentSlide == null)
            {
                MessageBox.Show(ShapesLabText.ErrorViewTypeNotSupported);
                return;
            }
            ClipboardUtil.RestoreClipboardAfterAction(() =>
            {
                // all selected shape will be added to the slide
                //TODO
                //Globals.ThisAddIn.ShapePresentation.CopyShape(_selectedThumbnail.Select(thumbnail => thumbnail.NameLabel));
                return currentSlide.Shapes.Paste();
            }, pres, currentSlide);
        }

        private void ContextMenuStripAddCategoryClicked()
        {
            ShapesLabCategoryInfoDialogBox categoryInfoDialog = new ShapesLabCategoryInfoDialogBox(string.Empty);
            categoryInfoDialog.DialogConfirmedHandler += (string newCategoryName) =>
            {
                Globals.ThisAddIn.ShapePresentation.AddCategory(newCategoryName);

                _categoryBinding.Add(newCategoryName);

                categoryBox.SelectedIndex = _categoryBinding.Count - 1;
            };
            categoryInfoDialog.ShowDialog();

            shapeList.Focus();
        }

        private void ContextMenuStripEditClicked()
        {
            //TODO
            /*
            if (_selectedThumbnail == null)
            {
                MessageBox.Show(ShapesLabText.ErrorNoPanelSelected);
                return;
            }*/

            // dehighlight all thumbnails except the first one
            /*
            while (_selectedThumbnail.Count > 1)
            {
                _selectedThumbnail[1].DeHighlight();
                _selectedThumbnail.RemoveAt(1);
            }

            _selectedThumbnail[0].StartNameEdit();*/
        }

        private void ContextMenuStripImportCategoryClicked()
        {
            //TODO
            /*
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Filter = ImportLibraryFileDialogFilter,
                Multiselect = false,
                Title = ShapesLabText.ImportLibraryFileDialogTitle
            };

            flowlayoutContextMenuStrip.Hide();

            if (fileDialog.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }

            ImportShapes(fileDialog.FileName, true);

            MessageBox.Show(ShapesLabText.SuccessImport);
            */
        }

        private void ContextMenuStripImportShapesClicked()
        {
            //TODO
            /*
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Filter = ImportShapesFileDialogFilter,
                Multiselect = true,
                Title = ShapesLabText.ImportShapeFileDialogTitle
            };

            if (fileDialog.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }

            bool importSuccess = fileDialog.FileNames.Aggregate(true,
                                                               (current, fileName) =>
                                                               ImportShapes(fileName, false) && current);

            if (!importSuccess)
            {
                return;
            }

            MessageBox.Show(ShapesLabText.SuccessImport);
            */
        }

        private void ContextMenuStripRemoveCategoryClicked()
        {
            // remove the last category will not be entertained
            if (_categoryBinding.Count == 1)
            {
                MessageBox.Show(ShapesLabText.ErrorRemoveLastCategory);
                return;
            }

            int categoryIndex = categoryBox.SelectedIndex;
            string categoryName = _categoryBinding[categoryIndex].ToString();
            string categoryPath = Path.Combine(ShapeRootFolderPath, categoryName);
            bool isDefaultCategory = Globals.ThisAddIn.ShapesLabConfig.DefaultCategory == CurrentCategory;

            if (isDefaultCategory)
            {
                DialogResult result =
                    MessageBox.Show(ShapesLabText.RemoveDefaultCategoryMessage,
                                    ShapesLabText.RemoveDefaultCategoryCaption,
                                    MessageBoxButtons.OKCancel);

                if (result == DialogResult.Cancel)
                {
                    return;
                }
            }

            // remove current category in shape gallery
            Globals.ThisAddIn.ShapePresentation.RemoveCategory();
            // remove category on the disk
            FileDir.DeleteFolder(categoryPath);

            _categoryBinding.RemoveAt(categoryIndex);

            // RemoveAt may NOT change the index, so we need to manually set the default category here
            if (Globals.ThisAddIn.ShapePresentation.DefaultCategory == null)
            {
                categoryIndex = categoryBox.SelectedIndex;
                categoryName = _categoryBinding[categoryIndex].ToString();

                CurrentCategory = categoryName;
                Globals.ThisAddIn.ShapePresentation.DefaultCategory = categoryName;
            }

            if (isDefaultCategory)
            {
                Globals.ThisAddIn.ShapesLabConfig.DefaultCategory = (string)_categoryBinding[0];
            }
        }

        private void ContextMenuStripRemoveClicked()
        {
            /*
            if (_selectedThumbnail == null ||
                _selectedThumbnail.Count == 0)
            {
                MessageBox.Show(ShapesLabText.ErrorNoPanelSelected);
                return;
            }*/

            //TODO
            //GraphicsUtil.SuspendDrawing(wrapPanel);

            /*
            while (_selectedThumbnail.Count > 0)
            {
                LabeledThumbnail thumbnail = _selectedThumbnail[0];
                string removedShapename = thumbnail.NameLabel;

                // remove shape from shape gallery
                Globals.ThisAddIn.ShapePresentation.RemoveShape(CurrentShapeNameWithoutExtension);

                // remove shape from disk and shape gallery
                File.Delete(CurrentShapeFullName);

                // remove shape from task pane
                RemoveShapeItem(thumbnail, false);

                // sync shape removing among all task panes
                Globals.ThisAddIn.SyncShapeRemove(removedShapename, CurrentCategory);

                // remove from selected collection
                _selectedThumbnail.RemoveAt(0);
            }
            */

            //TODO
            //GraphicsUtil.ResumeDrawing(wrapPanel);
        }

        private void ContextMenuStripRenameCategoryClicked()
        {
            ShapesLabCategoryInfoDialogBox categoryInfoDialog = new ShapesLabCategoryInfoDialogBox(string.Empty);
            categoryInfoDialog.DialogConfirmedHandler += (string newCategoryName) =>
            {
                // if current category is the default category, change ShapeConfig
                if (Globals.ThisAddIn.ShapesLabConfig.DefaultCategory == CurrentCategory)
                {
                    Globals.ThisAddIn.ShapesLabConfig.DefaultCategory = newCategoryName;
                }

                // rename the category in ShapeGallery
                Globals.ThisAddIn.ShapePresentation.RenameCategory(newCategoryName);

                // rename the category on the disk
                string newPath = Path.Combine(ShapeRootFolderPath, newCategoryName);

                try
                {
                    Directory.Move(CurrentShapeFolderPath, newPath);
                }
                catch (Exception)
                {
                    // this may occur when the newCategoryName.tolower() == oldCategoryName.tolower()
                }

                // rename the category in combo box
                int categoryIndex = categoryBox.SelectedIndex;
                _categoryBinding[categoryIndex] = newCategoryName;

                // update current category reference
                CurrentCategory = newCategoryName;
            };
            categoryInfoDialog.ShowDialog();

            shapeList.Focus();
        }

        private void ContextMenuStripSetAsDefaultCategoryClicked()
        {
            Globals.ThisAddIn.ShapesLabConfig.DefaultCategory = CurrentCategory;

            //TODO
            //comboBox.Refresh();

            MessageBox.Show(string.Format(ShapesLabText.SuccessSetAsDefaultCategory, CurrentCategory));
        }

        private void ContextMenuStripSettingsClicked()
        {
            ShapesLabSettingsDialogBox settingsDialog = new ShapesLabSettingsDialogBox(ShapeRootFolderPath);
            settingsDialog.DialogConfirmedHandler += (string newSavePath) =>
            {
                if (!MigrateShapeFolder(ShapeRootFolderPath, newSavePath))
                {
                    return;
                }

                ShapesLabSettings.SaveFolderPath = newSavePath;

                MessageBox.Show(
                    string.Format(ShapesLabText.SuccessSaveLocationChanged, newSavePath),
                    ShapesLabText.SuccessSaveLocationChangedTitle, MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            };
            settingsDialog.ShowDialog();
        }

        #endregion

        #region Helper Functions

        private void DehighlightSelected()
        {
            shapeList.UnselectAll();
        }

        private void DisableAddShapesButton()
        {
            addShapeButton.IsEnabled = false;
            addShapeImage.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    Properties.Resources.AddToCustomShapesDisabled.GetHbitmap(),
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());
            //TODO
            /*
            addShapeButton.FlatAppearance.BorderColor = Color.LightGray;
            addShapeButton.BackColor = Color.LightGray;*/
        }

        private void EnableAddShapesButton()
        {
            addShapeButton.IsEnabled = true;
            addShapeImage.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    Properties.Resources.AddToCustomShapes.GetHbitmap(),
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());
            //TODO
            /*
            addShapeButton.FlatAppearance.BorderColor = Color.Black;
            addShapeButton.BackColor = SystemColors.Control;*/
        }


        private int GetShapeItemIndex(string shapeName)
        {
            for (int index = 0; index < shapeList.Items.Count; index++)
            {
                if ((shapeList.Items[index] as CustomShapePaneItem).Text == shapeName)
                {
                    return index;
                }
            }
            return -1;
        }

        /// <summary>
        /// Returns the index at which the shape should be inserted lexicographically
        /// </summary>
        private int GetShapeInsertionIndex(string shapeName)
        {
            for (int index = 0; index < shapeList.Items.Count; index++)
            {
                if ((shapeList.Items[index] as CustomShapePaneItem).Text.CompareTo(shapeName) >= 0)
                {
                    shapeList.Items.RemoveAt(index);
                    return index;
                }
            }
            return shapeList.Items.Count;
        }

        private void ShapeListClick()
        {
            DehighlightSelected();
            shapeList.Focus();
        }

        private void FocusSelected()
        {
            //TODO
            //shapeList.ScrollIntoView();
        }

        private void RenameShapeFile(string oldShapeName, string newShapeName)
        {
            int shapeIndex = GetShapeItemIndex(oldShapeName);
            CustomShapePaneItem shapeItem = (shapeList.Items[shapeIndex] as CustomShapePaneItem);
            if (shapeItem.Text == newShapeName)
            {
                return;
            }
            shapeItem.RenameShape(newShapeName);

            //TODO
            Globals.ThisAddIn.ShapePresentation.RenameShape(oldShapeName, newShapeName);
            Globals.ThisAddIn.SyncShapeRename(oldShapeName, newShapeName, CurrentCategory);
        }

        private void PaneReload()
        {
            // clear all and load shapes from folder
            shapeList.Items.Clear();
            PrepareShapes();

            // scroll the view to show the first item
            shapeList.ScrollIntoView(shapeList.Items[0]);
            shapeList.Focus();
        }

        #endregion

        #region Shape Storage

        private bool ImportShapes(string importFilePath, bool fromLibrary)
        {
            PowerPointShapeGalleryPresentation importShapeGallery = PrepareImportGallery(importFilePath, fromLibrary);

            try
            {
                if (!importShapeGallery.Open(withWindow: false, focus: false))
                {
                    MessageBox.Show(ShapesLabText.ErrorImportFile);
                }
                else if (importShapeGallery.Slides.Count == 0)
                {
                    MessageBox.Show(ShapesLabText.ErrorImportNoSlide);
                }
                else
                {
                    // if user tries to import shapes but the file contains multiple categories,
                    // stop processing and warn the user
                    if (!fromLibrary && importShapeGallery.Categories.Count > 1)
                    {
                        MessageBox.Show(
                            string.Format(ShapesLabText.ErrorImportSingleCategory,
                                          importShapeGallery.Name));
                        return false;
                    }

                    // copy all shapes in the import shape gallery to current shape gallery
                    if (fromLibrary)
                    {
                        ImportShapesFromLibrary(importShapeGallery);
                    }
                    else
                    {
                        ImportShapesFromSingleShape(importShapeGallery);
                    }
                }
            }
            catch (Exception e)
            {
                ErrorDialogBox.ShowDialog(TextCollection.CommonText.ErrorTitle, e.Message, e);

                return false;
            }
            finally
            {
                importShapeGallery.Close();

                // delete the import file copy
                FileDir.DeleteFile(Path.Combine(ShapeRootFolderPath, ImportFileNameNoExtension + ".pptlabsshapes"));
            }

            return true;
        }

        private void ImportShapesFromLibrary(PowerPointShapeGalleryPresentation importShapeGallery)
        {
            // Utilises deprecated classes as CustomShapePane does not utilise ActionFramework
            PowerPointSlide currentSlide = this.GetCurrentSlide();
            PowerPointPresentation pres = this.GetCurrentPresentation();

            ClipboardUtil.RestoreClipboardAfterAction(() =>
            {
                foreach (string importCategory in importShapeGallery.Categories)
                {
                    importShapeGallery.CopyCategory(importCategory);

                    Globals.ThisAddIn.ShapePresentation.AddCategory(importCategory, false, true);

                    _categoryBinding.Add(importCategory);
                }
                return ClipboardUtil.ClipboardRestoreSuccess;
            }, pres, currentSlide);
        }

        private void ImportShapesFromSingleShape(PowerPointShapeGalleryPresentation importShapeGallery)
        {
            ShapeRange shapeRange = importShapeGallery.Slides[0].Shapes.Range();

            if (shapeRange.Count < 1)
            {
                return;
            }

            string shapeName = shapeRange[1].Name;

            // Utilises deprecated classes as CustomShapePane does not utilise ActionFramework
            PowerPointSlide currentSlide = this.GetCurrentSlide();
            PowerPointPresentation pres = this.GetCurrentPresentation();

            ClipboardUtil.RestoreClipboardAfterAction(() =>
            {
                importShapeGallery.CopyShape(shapeName);
                shapeName = Globals.ThisAddIn.ShapePresentation.AddShape(pres, currentSlide, null, shapeName, fromClipBoard: true);
                string exportPath = Path.Combine(CurrentShapeFolderPath, shapeName + ".png");

                GraphicsUtil.ExportShape(shapeRange, exportPath);
                return ClipboardUtil.ClipboardRestoreSuccess;
            }, pres, currentSlide);

        }

        private bool MigrateShapeFolder(string oldPath, string newPath)
        {
            LoadingDialogBox loadingDialog = new LoadingDialogBox(ShapesLabText.MigratingDialogTitle,
                                                    ShapesLabText.MigratingDialogContent);
            loadingDialog.Show();

            // close the opening presentation
            if (Globals.ThisAddIn.ShapePresentation.Opened)
            {
                Globals.ThisAddIn.ShapePresentation.Close();
            }

            // migration only cares about if the folder has been copied to the new location entirely.
            if (!FileDir.CopyFolder(oldPath, newPath))
            {
                loadingDialog.Close();

                MessageBox.Show(ShapesLabText.ErrorMigration);

                return false;
            }

            // now we will try our best to delete the original folder, but this is not guaranteed
            // because some of the using files, such as some opening shapes, and the evil thumb.db
            if (!FileDir.DeleteFolder(oldPath))
            {
                MessageBox.Show(ShapesLabText.ErrorOriginalFolderDeletion);
            }

            ShapeRootFolderPath = newPath;

            // modify shape gallery presentation's path and name, then open it
            Globals.ThisAddIn.ShapePresentation.Path = newPath;
            Globals.ThisAddIn.ShapePresentation.Open(withWindow: false, focus: false);
            Globals.ThisAddIn.ShapePresentation.DefaultCategory = CurrentCategory;

            loadingDialog.Close();

            return true;
        }

        private void PrepareFolder()
        {
            if (!Directory.Exists(CurrentShapeFolderPath))
            {
                Directory.CreateDirectory(CurrentShapeFolderPath);
            }
        }

        private PowerPointShapeGalleryPresentation PrepareImportGallery(string importFilePath, bool fromCategory)
        {
            string importFileCopyPath = Path.Combine(ShapeRootFolderPath, ImportFileCopyName);

            // copy the file to the current shape root if the file is not under root 
            if (!File.Exists(importFileCopyPath))
            {
                File.Copy(importFilePath, importFileCopyPath);
            }

            // init the file as an imported file
            PowerPointShapeGalleryPresentation importShapeGallery = new PowerPointShapeGalleryPresentation(ShapeRootFolderPath,
                                                                            ImportFileNameNoExtension)
            {
                IsImportedFile = true,
                ImportToCategory = fromCategory ? string.Empty : CurrentCategory
            };

            return importShapeGallery;
        }

        private void PrepareShapes()
        {
            PrepareFolder();

            IOrderedEnumerable<string> shapes =
                Directory.EnumerateFiles(CurrentShapeFolderPath, "*.png")
                .OrderBy(item => item, _stringComparer);

            foreach (string shape in shapes)
            {
                string shapeName = Path.GetFileNameWithoutExtension(shape);

                if (shapeName == null)
                {
                    MessageBox.Show(ShapesLabText.ErrorFileNameInvalid);
                    continue;
                }

                AddCustomShape(shapeName, shape, false);
            }
            //TODO
            /*
            if (myShapeFlowLayout.Controls.Count == 0)
            {
                ShowNoShapeMessage();
            }*/

            DehighlightSelected();
        }

        #endregion

        #region Event Handlers

        private void WrapPanelLoaded(object sender, RoutedEventArgs e)
        {
            wrapPanel = sender as WrapPanel;
        }

        private void CategoryBoxSelectedIndexChanged(object sender, EventArgs e)
        {
            CategoryBoxSelectedIndexChanged();
        }

        private void CategoryBoxSelectedIndexChanged()
        {
            int selectedIndex = categoryBox.SelectedIndex;
            string selectedCategory = _categoryBinding[selectedIndex].ToString();

            CurrentCategory = selectedCategory;
            Globals.ThisAddIn.ShapePresentation.DefaultCategory = selectedCategory;
            PaneReload();
        }

        #endregion

        #region GUI Handles

        private void AddShapeButton_Click(object sender, EventArgs e)
        {
            Selection selection = this.GetCurrentSelection();
            ThisAddIn addIn = this.GetAddIn();

            AddShapeFromSelection(selection, addIn);
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