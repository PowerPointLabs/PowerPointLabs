using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Models;
using PowerPointLabs.ShortcutsLab;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;

using MenuItem = System.Windows.Controls.MenuItem;
using MessageBox = System.Windows.Forms.MessageBox;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;

namespace PowerPointLabs.ShapesLab.Views
{
    /// <summary>
    /// Interaction logic for CustomShapePaneWPF.xaml
    /// </summary>
    public partial class CustomShapePaneWPF : System.Windows.Controls.UserControl
    {
        private const string ImportLibraryFileDialogFilter =
            "PowerPointLabs Shapes File|*.pptlabsshapes;*.pptx";
        private const string ImportShapesFileDialogFilter =
            "PowerPointLabs Shape File|*.pptlabsshape;*.pptx";
        private const string ImportFileNameNoExtension = "import";
        private const string ImportFileCopyName = ImportFileNameNoExtension + ".pptx";

        private readonly Comparers.AtomicNumberStringCompare _stringComparer = new Comparers.AtomicNumberStringCompare();

        private BindingSource _categoryBinding;
        private ObservableCollection<MenuItem> _contextMenuCategoryBinding;
        private WrapPanel wrapPanel;

        # region Properties
        public ObservableCollection<string> Categories { get; private set; }

        public string CurrentCategory { get; set; }

        public string ShapeRootFolderPath { get; private set; }

        public string CurrentShapeFolderPath
        {
            get { return ShapeRootFolderPath + @"\" + CurrentCategory; }
        }

        private bool IsStorageSettingsGiven
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
            DataContext = this;

            addShapeImage.Source = CommonUtil.CreateBitmapSource(Properties.Resources.AddToCustomShapes);

            singleShapeDownloadLink.NavigateUri = new Uri(CommonText.SingleShapeDownloadUrl);
        }

        #endregion

        #region Init

        public void SetStorageSettings()
        {
            if (IsStorageSettingsGiven)
            {
                return;
            }
            ThisAddIn addIn = this.GetAddIn();
            addIn.InitializeShapesLabConfig();
            addIn.InitializeShapeGallery();

            ShapeRootFolderPath = ShapesLabSettings.SaveFolderPath;
            CurrentCategory = addIn.ShapesLabConfig.DefaultCategory;
            Categories = new ObservableCollection<string>(this.GetAddIn().ShapePresentation.Categories);
            SetupMoveShapeContextMenu();
            SetupCategoryBoxItems();
        }

        public void CustomShapePaneWPF_Loaded(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Tools.CustomTaskPane shapesLabPane = this.GetAddIn().GetActivePane(typeof(CustomShapePane));
            CustomShapePane customShapePane = shapesLabPane?.Control as CustomShapePane;

            if (customShapePane == null)
            {
                MessageBox.Show(ShapesLabText.ErrorShapePaneNotOpened);
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

        public void AddShapeFromSelection(Selection selection)
        {
            ThisAddIn addIn = this.GetAddIn();
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
            ShapesLabUtils.SyncShapeAdd(addIn, shapeName, shapePath, CurrentCategory);

            // finally, add the shape into the panel
            AddCustomShape(shapeName, shapePath, false);
        }

        /// <summary>
        /// Adds a shape lexicographically.
        /// </summary>
        public void AddCustomShape(string shapeName, string shapePath, bool isReadyForEdit)
        {
            DehighlightSelected();

            CustomShapePaneItem shapeItem = new CustomShapePaneItem(this, shapeName, shapePath, isReadyForEdit, _contextMenuCategoryBinding);

            int insertionIndex = GetShapeInsertionIndex(shapeName);
            shapeList.Items.Insert(insertionIndex, shapeItem);
            shapeList.SelectedIndex = insertionIndex;
            shapeList.ScrollIntoView(shapeItem);
        }

        public void RemoveCustomShape(string shapeName)
        {
            int shapeIndex = GetShapeItemIndex(shapeName);
            if (shapeIndex < 0)
            {
                return;
            }
            shapeList.Items.RemoveAt(shapeIndex);
        }

        public void RemoveAllSelectedShapes()
        {
            //store names in list first, as enumeration will fail if the selected items are modified.
            List<string> shapeNames = new List<string>();
            foreach (CustomShapePaneItem shape in shapeList.SelectedItems)
            {
                shapeNames.Add(shape.Text);
            }
            foreach (string shapeName in shapeNames)
            {
                RemoveShape(shapeName);
            }
        }

        public void RenameCustomShape(string oldShapeName, string newShapeName)
        {
            int shapeIndex = GetShapeItemIndex(oldShapeName);
            if (shapeIndex < 0)
            {
                return;
            }
            CustomShapePaneItem shape = shapeList.Items[shapeIndex] as CustomShapePaneItem;
            shape.SyncRenameShape(newShapeName);
            shapeList.Items.Remove(shape);
            int insertionIndex = GetShapeInsertionIndex(newShapeName);
            shapeList.Items.Insert(insertionIndex, shape);
            shapeList.SelectedIndex = insertionIndex;
            shapeList.ScrollIntoView(shape);
        }

        public bool IsShapeSelected(CustomShapePaneItem shape)
        {
            return shapeList.SelectedItems.Contains(shape);
        }

        public void AddShapesToSlide()
        {
            if (shapeList.SelectedItems.Count == 0)
            {
                MessageBox.Show(ShapesLabText.ErrorNoPanelSelected, ShapesLabText.ErrorDialogTitle);
                return;
            }
            PowerPointSlide currentSlide = this.GetCurrentSlide();
            if (currentSlide == null)
            {
                MessageBox.Show(ShapesLabText.ErrorViewTypeNotSupported);
                return;
            }
            this.StartNewUndoEntry();
            foreach (CustomShapePaneItem shape in shapeList.SelectedItems)
            {
                ClipboardUtil.RestoreClipboardAfterAction(() =>
                {
                    this.GetAddIn().ShapePresentation.CopyShape(shape.Text);
                    currentSlide.Shapes.Paste().Select();
                    return ClipboardUtil.ClipboardRestoreSuccess;
                }, this.GetCurrentPresentation(), currentSlide);
            }
        }

        public void MoveShapes(string newCategoryName)
        {
            //Add shape names first as shapeList.Items will be modified
            List<string> shapeNames = new List<string>();
            foreach (CustomShapePaneItem shape in shapeList.SelectedItems)
            {
                shapeNames.Add(shape.Text);
            }

            foreach (string shapeName in shapeNames)
            {
                string oriPath = Path.Combine(CurrentShapeFolderPath, shapeName) + ".png";
                string destPath = Path.Combine(ShapeRootFolderPath, newCategoryName, shapeName) + ".png";

                // if we have an identical name in the destination category, we won't allow
                // moving
                if (File.Exists(destPath))
                {
                    MessageBox.Show(string.Format(TextCollection.ShapesLabText.ErrorSameShapeNameInDestination,
                                    shapeName,
                                    newCategoryName));
                    return;
                }
                PowerPointSlide currentSlide = this.GetCurrentSlide();
                PowerPointPresentation pres = this.GetCurrentPresentation();

                // move shape in ShapeGallery to correct place
                Globals.ThisAddIn.ShapePresentation.MoveShapeToCategory(pres, currentSlide, shapeName, newCategoryName);
                File.Move(oriPath, destPath);
                RemoveCustomShape(shapeName);

                ShapesLabUtils.SyncShapeRemove(this.GetAddIn(), shapeName, CurrentCategory);
                ShapesLabUtils.SyncShapeAdd(this.GetAddIn(), shapeName, destPath, newCategoryName);
            }
        }

        public void AddCategory(string newCategoryName)
        {
            this.GetAddIn().ShapePresentation.AddCategory(newCategoryName);
            _categoryBinding.Add(new CustomComboBoxItem(newCategoryName, null));
            _contextMenuCategoryBinding.Add(new CustomMenuItem(newCategoryName, MoveShapeClick));
        }

        public void RemoveCategory(int removedCategoryIndex)
        {
            int categoryIndex = categoryBox.SelectedIndex;
            _categoryBinding.RemoveAt(categoryIndex);
            _contextMenuCategoryBinding.RemoveAt(categoryIndex);
            if (categoryIndex == removedCategoryIndex)
            {
                categoryBox.SelectedIndex = Math.Max(0, categoryIndex - 1);
            }
        }

        public void RenameCategory(int renameCategoryIndex, string newCategoryName)
        {
            bool isCurrentCategoryRenamed = renameCategoryIndex == categoryBox.SelectedIndex;
            CustomComboBoxItem item = _categoryBinding[renameCategoryIndex] as CustomComboBoxItem;
            if (item == null)
            {
                return;
            }
            item.actualName = newCategoryName;
            _categoryBinding.RemoveAt(renameCategoryIndex);
            _categoryBinding.Insert(renameCategoryIndex, item);
            _contextMenuCategoryBinding.RemoveAt(renameCategoryIndex);
            _contextMenuCategoryBinding.Insert(renameCategoryIndex, new CustomMenuItem(newCategoryName, MoveShapeClick));
            if (isCurrentCategoryRenamed)
            {
                CurrentCategory = newCategoryName;
                categoryBox.SelectedIndex = renameCategoryIndex;
            }
        }

        #endregion

        #region Test Interface

        public CustomShapePaneItem GetShape(string shapeName)
        {
            int shapeIndex = GetShapeItemIndex(shapeName);
            if (shapeIndex == -1)
            {
                return null;
            }
            return (CustomShapePaneItem) shapeList.Items[shapeIndex];
        }

        public void ImportLibrary(string pathToLibrary)
        {
            ImportShapes(pathToLibrary, true);
        }

        public void ImportShape(string pathToShape)
        {
            ImportShapes(pathToShape, false);
        }

        public Presentation GetShapeGallery()
        {
            return this.GetAddIn().ShapePresentation.Presentation;
        }

        public void SaveSelectedShapes()
        {
            Selection selection = this.GetCurrentSelection();
            AddShapeFromSelection(selection);
        }

        public System.Windows.Point GetShapeForClicking(string shapeName)
        {
            int shapeIndex = GetShapeItemIndex(shapeName);
            if (shapeIndex == -1)
            {
                return new System.Windows.Point(0, 0);
            }
            CustomShapePaneItem shape = (CustomShapePaneItem)shapeList.Items[shapeIndex];
            shape.FinishNameEdit();
            return shape.grid.PointToScreen(new System.Windows.Point(20, 20));
        }

        #endregion

        #region Context Menu

        private void ContextMenuStripAddCategoryClicked(object sender, RoutedEventArgs e)
        {
            ShapesLabCategoryInfoDialogBox categoryInfoDialog = new ShapesLabCategoryInfoDialogBox(string.Empty, false);
            categoryInfoDialog.DialogConfirmedHandler += (string newCategoryName) =>
            {
                ShapesLabUtils.SyncAddCategory(this.GetAddIn(), newCategoryName);
                AddCategory(newCategoryName);
                categoryBox.SelectedIndex = _categoryBinding.Count - 1;
            };
            categoryInfoDialog.ShowDialog();
            shapeList.Focus();
        }

        private void ContextMenuStripImportCategoryClicked(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Filter = ImportLibraryFileDialogFilter,
                Multiselect = false,
                Title = ShapesLabText.ImportLibraryFileDialogTitle
            };

            //flowlayoutContextMenuStrip.Hide();

            if (fileDialog.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }

            ImportShapes(fileDialog.FileName, true);

            MessageBox.Show(ShapesLabText.SuccessImport);
        }

        private void ContextMenuStripImportShapesClicked(object sender, RoutedEventArgs e)
        {
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
        }

        private void ContextMenuStripRemoveCategoryClicked(object sender, RoutedEventArgs e)
        {
            // remove the last category will not be entertained
            if (_categoryBinding.Count == 1)
            {
                MessageBox.Show(ShapesLabText.ErrorRemoveLastCategory);
                return;
            }

            int categoryIndex = categoryBox.SelectedIndex;
            string categoryName = (_categoryBinding[categoryIndex] as CustomComboBoxItem)?.actualName;
            string categoryPath = Path.Combine(ShapeRootFolderPath, categoryName);
            bool isDefaultCategory = this.GetAddIn().ShapesLabConfig.DefaultCategory == CurrentCategory;

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
            this.GetAddIn().ShapePresentation.RemoveCategory();
            // remove category on the disk
            FileDir.DeleteFolder(categoryPath);

            ShapesLabUtils.SyncRemoveCategory(this.GetAddIn(), categoryIndex);
            RemoveCategory(categoryIndex);

            if (this.GetAddIn().ShapePresentation.DefaultCategory == null)
            {
                this.GetAddIn().ShapePresentation.DefaultCategory = CurrentCategory;
            }

            if (isDefaultCategory)
            {
                CustomComboBoxItem item = _categoryBinding[0] as CustomComboBoxItem;
                this.GetAddIn().ShapesLabConfig.DefaultCategory = item?.actualName;
                item?.SetNewDefaultCategory(this.GetAddIn().ShapesLabConfig.DefaultCategory);
                _categoryBinding.RemoveAt(0);
                _categoryBinding.Insert(0, item);
                categoryBox.SelectedIndex = 0;
            }
        }

        private void ContextMenuStripRenameCategoryClicked(object sender, RoutedEventArgs e)
        {
            ShapesLabCategoryInfoDialogBox categoryInfoDialog = new ShapesLabCategoryInfoDialogBox(string.Empty, false);
            categoryInfoDialog.DialogConfirmedHandler += (string newCategoryName) =>
            {
                // if current category is the default category, change ShapeConfig
                if (this.GetAddIn().ShapesLabConfig.DefaultCategory == CurrentCategory)
                {
                    this.GetAddIn().ShapesLabConfig.DefaultCategory = newCategoryName;
                }

                // rename the category in ShapeGallery
                this.GetAddIn().ShapePresentation.RenameCategory(newCategoryName);

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

                ShapesLabUtils.SyncRenameCategory(this.GetAddIn(), categoryIndex, newCategoryName);
                RenameCategory(categoryIndex, newCategoryName);
            };
            categoryInfoDialog.ShowDialog();

            shapeList.Focus();
        }

        private void ContextMenuStripSetAsDefaultCategoryClicked(object sender, RoutedEventArgs e)
        {
            this.GetAddIn().ShapesLabConfig.DefaultCategory = CurrentCategory;
            MessageBox.Show(string.Format(ShapesLabText.SuccessSetAsDefaultCategory, CurrentCategory));
            for (int index = 0; index < _categoryBinding.Count; index++)
            {
                CustomComboBoxItem category = _categoryBinding[index] as CustomComboBoxItem;
                if (category == null || (!category.isDefaultCategory && category.actualName != CurrentCategory))
                {
                    continue;
                }
                category.SetNewDefaultCategory(CurrentCategory);
                _categoryBinding.RemoveAt(index);
                _categoryBinding.Insert(index, category);
                if (category.actualName == CurrentCategory)
                {
                    categoryBox.SelectedIndex = index;
                }
            }
        }

        private void ContextMenuStripSettingsClicked(object sender, RoutedEventArgs e)
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

        private void SetupCategoryBoxItems()
        {
            ObservableCollection<CustomComboBoxItem> categoryBoxItems = new ObservableCollection<CustomComboBoxItem>();
            string defaultCategory = this.GetAddIn().ShapePresentation.DefaultCategory;
            foreach (string category in Categories)
            {
                categoryBoxItems.Add(new CustomComboBoxItem(category, defaultCategory));
            }
            _categoryBinding = new BindingSource { DataSource = categoryBoxItems };
            categoryBox.ItemsSource = _categoryBinding;

            for (int i = 0; i < Categories.Count; i++)
            {
                if (Categories[i] == CurrentCategory)
                {
                    categoryBox.SelectedIndex = i;
                    break;
                }
            }
        }

        private void SetupMoveShapeContextMenu()
        {
            _contextMenuCategoryBinding = new ObservableCollection<MenuItem>();
            foreach (string category in Categories)
            {
                _contextMenuCategoryBinding.Add(new CustomMenuItem(category, MoveShapeClick));
            }
        }

        private void DehighlightSelected()
        {
            shapeList.UnselectAll();
        }

        private void DisableAddShapesButton()
        {
            addShapeButton.IsEnabled = false;
            addShapeImage.Source = CommonUtil.CreateBitmapSource(Properties.Resources.AddToCustomShapesDisabled);
        }

        private void EnableAddShapesButton()
        {
            addShapeButton.IsEnabled = true;
            addShapeImage.Source = CommonUtil.CreateBitmapSource(Properties.Resources.AddToCustomShapes);
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
                    return index;
                }
            }
            return shapeList.Items.Count;
        }

        private void PaneReload()
        {
            // clear all and load shapes from folder
            shapeList.Items.Clear();
            PrepareShapes();

            if (shapeList.Items.IsEmpty)
            {
                return;
            }
            // scroll the view to show the first item
            shapeList.ScrollIntoView(shapeList.Items[0]);
            shapeList.Focus();
        }

        private string GetShapePath(string shapeName)
        {
            return CurrentShapeFolderPath + @"\" + shapeName + ".png";
        }

        private void RemoveShape(string shapeName)
        {
            string shapePath = GetShapePath(shapeName);
            if (!File.Exists(shapePath))
            {
                return;
            }
            // remove shape from disk and shape gallery
            File.Delete(shapePath);

            // remove shape from shape gallery
            this.GetAddIn().ShapePresentation.RemoveShape(shapeName);

            // sync shape removing among all task panes
            ShapesLabUtils.SyncShapeRemove(this.GetAddIn(), shapeName, CurrentCategory);
            RemoveCustomShape(shapeName);
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
                ErrorDialogBox.ShowDialog(CommonText.ErrorTitle, e.Message, e);

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
            PowerPointSlide currentSlide = this.GetCurrentSlide();
            PowerPointPresentation pres = this.GetCurrentPresentation();

            ClipboardUtil.RestoreClipboardAfterAction(() =>
            {
                foreach (string importCategory in importShapeGallery.Categories)
                {
                    importShapeGallery.CopyCategory(importCategory);

                    this.GetAddIn().ShapePresentation.AddCategory(importCategory, false, true);

                    _categoryBinding.Add(new CustomComboBoxItem(importCategory, null));
                    _contextMenuCategoryBinding.Add(new CustomMenuItem(importCategory, MoveShapeClick));
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
                shapeName = this.GetAddIn().ShapePresentation.AddShape(pres, currentSlide, null, shapeName, fromClipBoard: true);
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
            if (this.GetAddIn().ShapePresentation.Opened)
            {
                this.GetAddIn().ShapePresentation.Close();
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
            this.GetAddIn().ShapePresentation.Path = newPath;
            this.GetAddIn().ShapePresentation.Open(withWindow: false, focus: false);
            this.GetAddIn().ShapePresentation.DefaultCategory = CurrentCategory;

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
            int selectedIndex = categoryBox.SelectedIndex;
            if (selectedIndex == -1)
            {
                return;
            }
            string selectedCategory = (_categoryBinding[selectedIndex] as CustomComboBoxItem)?.actualName;

            CurrentCategory = selectedCategory;
            this.GetAddIn().ShapePresentation.DefaultCategory = selectedCategory;
            PaneReload();
        }

        private void AddShapeButton_Click(object sender, EventArgs e)
        {
            Selection selection = this.GetCurrentSelection();

            AddShapeFromSelection(selection);
        }

        private void ClickOutsideShapeList(object sender, MouseButtonEventArgs e)
        {
            categoryBox.Focus();
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }

        private void MoveShapeClick(object sender, RoutedEventArgs e)
        {
            CustomMenuItem item = sender as CustomMenuItem;
            if (item == null)
            {
                return;
            }
            MoveShapes(item.actualName);
        }

        #endregion
    }
}