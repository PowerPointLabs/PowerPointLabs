using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
        private WrapPanel wrapPanel;

        # region Properties
        public ObservableCollection<string> Categories { get; private set; }

        public string CurrentCategory { get; set; }

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
            DataContext = this;

            addShapeImage.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    Properties.Resources.AddToCustomShapes.GetHbitmap(),
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());

            //singleShapeDownloadLink.LinkClicked += (s, e) => Process.Start(CommonText.SingleShapeDownloadUrl);

        }

        #endregion

        #region Init

        public void SetStorageSettings(string shapeRootFolderPath, string defaultShapeCategoryName)
        {
            ShapeRootFolderPath = shapeRootFolderPath;

            CurrentCategory = defaultShapeCategoryName;
            Categories = new ObservableCollection<string>(Globals.ThisAddIn.ShapePresentation.Categories);
            _categoryBinding = new BindingSource { DataSource = Categories };
            categoryBox.ItemsSource = _categoryBinding;

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
        public void AddCustomShape(string shapeName, string shapePath, bool isReadyForEdit)
        {
            DehighlightSelected();

            CustomShapePaneItem shapeItem = new CustomShapePaneItem(this, shapeName, shapePath, isReadyForEdit);

            //shapeItem.Image = new System.Drawing.Bitmap(GraphicsUtil.ShapeToBitmap(shape));
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
                    Globals.ThisAddIn.ShapePresentation.CopyShape(shape.Text);
                    currentSlide.Shapes.Paste().Select();
                    return ClipboardUtil.ClipboardRestoreSuccess;
                }, this.GetCurrentPresentation(), currentSlide);
            }
        }

        #endregion

        #region Context Menu
        
        private void ContextMenuStripAddCategoryClicked(object sender, RoutedEventArgs e)
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

        private void ContextMenuStripRenameCategoryClicked(object sender, RoutedEventArgs e)
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

        private void ContextMenuStripSetAsDefaultCategoryClicked(object sender, RoutedEventArgs e)
        {
            Globals.ThisAddIn.ShapesLabConfig.DefaultCategory = CurrentCategory;

            //TODO
            //comboBox.Refresh();

            MessageBox.Show(string.Format(ShapesLabText.SuccessSetAsDefaultCategory, CurrentCategory));
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
            this.GetAddIn().SyncShapeRemove(shapeName, CurrentCategory);
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

            DehighlightSelected();
        }

        private void CategoryChanged(object sender, SelectionChangedEventArgs e)
        {

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
            this.GetAddIn().ShapePresentation.DefaultCategory = selectedCategory;
            PaneReload();
        }

        private void AddShapeButton_Click(object sender, EventArgs e)
        {
            Selection selection = this.GetCurrentSelection();
            ThisAddIn addIn = this.GetAddIn();

            AddShapeFromSelection(selection, addIn);
        }

        #endregion

    }
}