using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Models;
using PowerPointLabs.ShapesLab.Views;
using PowerPointLabs.ShortcutsLab;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;

using PPExtraEventHelper;

using Font = System.Drawing.Font;
using GraphicsUtil = PowerPointLabs.Utils.GraphicsUtil;
using Point = System.Drawing.Point;

namespace PowerPointLabs.ShapesLab
{
    public partial class CustomShapePane_ : UserControl
    {
#pragma warning disable 0618
        private const string ImportLibraryFileDialogFilter =
            "PowerPointLabs Shapes File|*.pptlabsshapes;*.pptx";
        private const string ImportShapesFileDialogFilter =
            "PowerPointLabs Shape File|*.pptlabsshape;*.pptx";
        private const string ImportFileNameNoExtension = "import";
        private const string ImportFileCopyName = ImportFileNameNoExtension + ".pptx";

        private readonly int _doubleClickTimeSpan = SystemInformation.DoubleClickTime;
        private int _clicks;

        private bool _firstTimeLoading = true;
        private bool _firstClick = true;
        private bool _clickOnSelected;
        private bool _isLeftButton;
        private bool _toolTipShown = false;

        private bool _isPanelMouseDown;
        private bool _isPanelDrawingFinish;
        private Point _startPosition;
        private Point _curPosition;

        private readonly SelectionRectangle _selectRect = new SelectionRectangle();

        private readonly BindingSource _categoryBinding;

        private List<LabeledThumbnail> _selectedThumbnail = new List<LabeledThumbnail>();
        private List<LabeledThumbnail> _selectingThumbnail = new List<LabeledThumbnail>();

        private readonly Timer _timer;

        private readonly Comparers.AtomicNumberStringCompare _stringComparer = new Comparers.AtomicNumberStringCompare();

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

                if (myShapeFlowLayout.Controls.Count == 0 ||
                    myShapeFlowLayout.Controls.Contains(_noShapePanel))
                {
                    return shapeList;
                }

                shapeList.AddRange(from LabeledThumbnail labelThumbnail in myShapeFlowLayout.Controls
                                   select labelThumbnail.NameLabel);

                return shapeList;
            }
        }

        public string ShapeRootFolderPath { get; private set; }

        public string CurrentShapeFolderPath
        {
            get { return ShapeRootFolderPath + @"\" + CurrentCategory; }
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams createParams = base.CreateParams;

                // do this optimization only for office 2010 since painting speed on 2013 is
                // really slow
                if (Globals.ThisAddIn.IsApplicationVersion2010())
                {
                    createParams.ExStyle |= (int)Native.Message.WS_EX_COMPOSITED;  // Turn on WS_EX_COMPOSITED
                }

                return createParams;
            }
        }
        # endregion

        # region Constructors
        public CustomShapePane_(string shapeRootFolderPath, string defaultShapeCategoryName)
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.DoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
            InitializeComponent();
            
            InitToolTipControl();

            InitializeContextMenu();

            ShapeRootFolderPath = shapeRootFolderPath;

            CurrentCategory = defaultShapeCategoryName;
            Categories = new List<string>(Globals.ThisAddIn.ShapePresentation.Categories);
            _categoryBinding = new BindingSource { DataSource = Categories };
            categoryBox.DataSource = _categoryBinding;

            for (int i = 0; i < Categories.Count; i++)
            {
                if (Categories[i] == defaultShapeCategoryName)
                {
                    categoryBox.SelectedIndex = i;
                    break;
                }
            }

            _timer = new Timer { Interval = _doubleClickTimeSpan };
            _timer.Tick += TimerTickHandler;

            myShapeFlowLayout.AutoSize = true;
            myShapeFlowLayout.MouseEnter += FlowLayoutMouseEnterHandler;
            myShapeFlowLayout.MouseDown += FlowLayoutMouseDownHandler;
            myShapeFlowLayout.MouseUp += FlowLayoutMouseUpHandler;
            myShapeFlowLayout.MouseMove += FlowLayoutMouseMoveHandler;

            singleShapeDownloadLink.LinkClicked += (s, e) => Process.Start(CommonText.SingleShapeDownloadUrl);
        }
        #endregion

        #region API

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

            // Utilises deprecated classes as CustomShapePane does not utilise ActionFramework
            PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
            PowerPointPresentation pres = PowerPointPresentation.Current;

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
            DehighlightSelected();

            LabeledThumbnail labeledThumbnail = new LabeledThumbnail(shapePath, shapeName) { ContextMenuStrip = shapeContextMenuStrip };

            labeledThumbnail.Click += LabeledThumbnailClick;
            labeledThumbnail.DoubleClick += LabeledThumbnailDoubleClick;
            labeledThumbnail.NameEditFinish += NameEditFinishHandler;

            myShapeFlowLayout.Controls.Add(labeledThumbnail);

            if (myShapeFlowLayout.Controls.Contains(_noShapePanel))
            {
                myShapeFlowLayout.Controls.Remove(_noShapePanel);
            }

            myShapeFlowLayout.ScrollControlIntoView(labeledThumbnail);

            _selectedThumbnail.Insert(0, labeledThumbnail);

            if (immediateEditing)
            {
                labeledThumbnail.StartNameEdit();
            }
            else
            {
                // the shape must be sorted immediately since the name has been
                // setteled
                ReorderThumbnail(labeledThumbnail);
            }
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

            // renaming will possibly change the relative shape order, therefore we need
            // to reorder the labeled thumbnail
            ReorderThumbnail(labeledThumbnail);

            // highlight the thumbnail again in case it is the selected shape
            if (labeledThumbnail == _selectedThumbnail[0])
            {
                labeledThumbnail.Highlight();
            }
        }

        public void PaneReload(bool forceReload = false)
        {
            if (!_firstTimeLoading && !forceReload)
            {
                return;
            }

            // double buffer starts
            if (Globals.ThisAddIn.IsApplicationVersion2013())
            {
                GraphicsUtil.SuspendDrawing(myShapeFlowLayout);
            }

            // emptize the panel and load shapes from folder
            myShapeFlowLayout.Controls.Clear();
            PrepareShapes();

            // scroll the view to show the first item, and focus the flowlayout to enable
            // scroll if applicable
            myShapeFlowLayout.ScrollControlIntoView(myShapeFlowLayout.Controls[0]);
            myShapeFlowLayout.Focus();

            // double buffer ends
            if (Globals.ThisAddIn.IsApplicationVersion2013())
            {
                GraphicsUtil.ResumeDrawing(myShapeFlowLayout);
            }

            _firstTimeLoading = false;
        }

        public void UpdateOnSelectionChange(Selection selection)
        {
            SelectionChanged(selection);
        }
        #endregion

        #region Functional Test APIs

        public LabeledThumbnail GetLabeledThumbnail(string labelName)
        {
            return FindLabeledThumbnail(labelName);
        }

        public void ImportLibrary(string pathToLibrary)
        {
            ImportShapes(pathToLibrary, fromLibrary: true);
        }

        public void ImportShape(string pathToShape)
        {
            ImportShapes(pathToShape, fromLibrary: false);
        }

        public Presentation GetShapeGallery()
        {
            return Globals.ThisAddIn.ShapePresentation.Presentation;
        }

        public Button GetAddShapeButton()
        {
            return addShapeButton;
        }

        # endregion

        # region Helper Functions
        private void ClickTimerReset()
        {
            _clicks = 0;
            _clickOnSelected = false;
            _firstClick = true;
            _isLeftButton = false;
        }

        private void ContextMenuStripAddClicked()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;

            // Utilises deprecated PowerPointPresentation class as CustomShapePane does not utilise ActionFramework
            PowerPointPresentation pres = PowerPointPresentation.Current;

            if (currentSlide == null)
            {
                MessageBox.Show(ShapesLabText.ErrorViewTypeNotSupported);
                return;
            }
            ClipboardUtil.RestoreClipboardAfterAction(() =>
            {
                // all selected shape will be added to the slide
                Globals.ThisAddIn.ShapePresentation.CopyShape(_selectedThumbnail.Select(thumbnail => thumbnail.NameLabel));
                return currentSlide.Shapes.Paste();
            }, pres, currentSlide);
        }

        private void ContextMenuStripAddCategoryClicked()
        {
            ShapesLabCategoryInfoDialogBox categoryInfoDialog = new ShapesLabCategoryInfoDialogBox(string.Empty, false);
            categoryInfoDialog.DialogConfirmedHandler += (string newCategoryName) =>
            {
                Globals.ThisAddIn.ShapePresentation.AddCategory(newCategoryName);

                _categoryBinding.Add(newCategoryName);

                categoryBox.SelectedIndex = _categoryBinding.Count - 1;
            };
            categoryInfoDialog.ShowDialog();

            myShapeFlowLayout.Focus();
        }

        private void ContextMenuStripEditClicked()
        {
            if (_selectedThumbnail == null)
            {
                MessageBox.Show(ShapesLabText.ErrorNoPanelSelected);
                return;
            }

            // dehighlight all thumbnails except the first one
            while (_selectedThumbnail.Count > 1)
            {
                _selectedThumbnail[1].DeHighlight();
                _selectedThumbnail.RemoveAt(1);
            }

            _selectedThumbnail[0].StartNameEdit();
        }

        private void ContextMenuStripImportCategoryClicked()
        {
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
        }

        private void ContextMenuStripImportShapesClicked()
        {
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Filter = ImportShapesFileDialogFilter,
                Multiselect = true,
                Title = ShapesLabText.ImportShapeFileDialogTitle
            };

            flowlayoutContextMenuStrip.Hide();

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

            PaneReload(true);
            MessageBox.Show(ShapesLabText.SuccessImport);
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

                PaneReload(true);
            }

            if (isDefaultCategory)
            {
                Globals.ThisAddIn.ShapesLabConfig.DefaultCategory = (string)_categoryBinding[0];
            }
        }

        private void ContextMenuStripRemoveClicked()
        {
            if (_selectedThumbnail == null ||
                _selectedThumbnail.Count == 0)
            {
                MessageBox.Show(ShapesLabText.ErrorNoPanelSelected);
                return;
            }

            GraphicsUtil.SuspendDrawing(myShapeFlowLayout);

            while (_selectedThumbnail.Count > 0)
            {
                LabeledThumbnail thumbnail = _selectedThumbnail[0];
                string removedShapename = thumbnail.NameLabel;

                // remove shape from shape gallery
                Globals.ThisAddIn.ShapePresentation.RemoveShape(CurrentShapeNameWithoutExtension);

                // remove shape from disk and shape gallery
                File.Delete(CurrentShapeFullName);

                // remove shape from task pane
                RemoveThumbnail(thumbnail, false);

                // sync shape removing among all task panes
                Globals.ThisAddIn.SyncShapeRemove(removedShapename, CurrentCategory);

                // remove from selected collection
                _selectedThumbnail.RemoveAt(0);
            }

            GraphicsUtil.ResumeDrawing(myShapeFlowLayout);
        }

        private void ContextMenuStripRenameCategoryClicked()
        {
            ShapesLabCategoryInfoDialogBox categoryInfoDialog = new ShapesLabCategoryInfoDialogBox(string.Empty, false);
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

            myShapeFlowLayout.Focus();
        }

        private void ContextMenuStripSetAsDefaultCategoryClicked()
        {
            Globals.ThisAddIn.ShapesLabConfig.DefaultCategory = CurrentCategory;

            categoryBox.Refresh();

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

        private Rectangle CreateRect(Point loc1, Point loc2)
        {
            RegulateSelectionRectPoint(ref loc1);
            RegulateSelectionRectPoint(ref loc2);

            Size size = new Size(Math.Abs(loc2.X - loc1.X), Math.Abs(loc2.Y - loc1.Y));
            Rectangle rect = new Rectangle(new Point(Math.Min(loc1.X, loc2.X), Math.Min(loc1.Y, loc2.Y)), size);

            return rect;
        }

        private void DehighlightSelected()
        {
            if (_selectedThumbnail == null ||
                _selectedThumbnail.Count == 0)
            {
                return;
            }

            foreach (LabeledThumbnail thumbnail in _selectedThumbnail)
            {
                thumbnail.DeHighlight();
            }

            _selectedThumbnail.Clear();
        }

        private void DisableAddShapesButton()
        {
            addShapeButton.Enabled = false;
            addShapeButton.BackgroundImage = Properties.Resources.AddToCustomShapesDisabled;
            addShapeButton.FlatAppearance.BorderColor = Color.LightGray;
            addShapeButton.BackColor = Color.LightGray;
        }

        private void EnableAddShapesButton()
        {
            addShapeButton.Enabled = true;
            addShapeButton.BackgroundImage = Properties.Resources.AddToCustomShapes;
            addShapeButton.FlatAppearance.BorderColor = Color.Black;
            addShapeButton.BackColor = SystemColors.Control;
        }


        private LabeledThumbnail FindLabeledThumbnail(string name)
        {
            if (myShapeFlowLayout.Controls.Count == 0 ||
                myShapeFlowLayout.Controls.Contains(_noShapePanel))
            {
                return null;
            }

            return
                myShapeFlowLayout.Controls.Cast<LabeledThumbnail>().FirstOrDefault(
                    labeledThumbnail => labeledThumbnail.NameLabel == name);
        }

        private int FindLabeledThumbnailIndex(string name)
        {
            if (myShapeFlowLayout.Controls.Contains(_noShapePanel))
            {
                return -1;
            }

            int totalControl = myShapeFlowLayout.Controls.Count;
            int thisControlPosition = -1;

            for (int i = 0; i < totalControl; i++)
            {
                LabeledThumbnail control = myShapeFlowLayout.Controls[i] as LabeledThumbnail;

                if (control == null)
                {
                    continue;
                }

                // skip itself
                if (control.NameLabel == name)
                {
                    thisControlPosition = i;
                    continue;
                }

                if (_stringComparer.Compare(control.NameLabel, name) > 0)
                {
                    // immediate next control's name is still bigger than current control, do
                    // not need to reorder
                    if (thisControlPosition != -1 &&
                        i - 1 == thisControlPosition)
                    {
                        return thisControlPosition;
                    }

                    // now we have 2 cases:
                    // 1. the replace position is before the current position;
                    // 2. the replace position is behind the current position.
                    // For case 1, we just need to set the current control's index to replace
                    // position, Windows Form will handle the rest of control's order;
                    // For case 2, we need to set the current control's index to replace position - 1,
                    // this is because the shapes behind will move forward by 1 when the current
                    // shape is moved.
                    if (thisControlPosition == -1)
                    {
                        // case 1, we haven't encountered the current control yet
                        return i;
                    }

                    // case 2, we have encountered the current control
                    return i - 1;
                }
            }

            return totalControl - 1;
        }

        private void FirstClickOnThumbnail(LabeledThumbnail clickedThumbnail)
        {
            if (_selectedThumbnail == null)
            {
                return;
            }

            if (_selectedThumbnail.Count != 0)
            {
                // this flag doesn't apply for multi selection, thus turn off
                _clickOnSelected = false;

                // common part, end editing
                if (_selectedThumbnail[0].State == LabeledThumbnail.Status.Editing)
                {
                    _selectedThumbnail[0].FinishNameEdit();
                }
                else
                if (_selectedThumbnail[0] == clickedThumbnail)
                {
                    _clickOnSelected = true;
                }

                MultiSelectClickHandler(clickedThumbnail);
            }
            else
            {
                clickedThumbnail.Highlight();

                _selectedThumbnail.Insert(0, clickedThumbnail);
                FocusSelected();
            }
        }

        private void FlowlayoutClick()
        {
            if (_selectedThumbnail != null &&
                _selectedThumbnail.Count != 0)
            {
                if (_selectedThumbnail[0].State == LabeledThumbnail.Status.Editing)
                {
                    _selectedThumbnail[0].FinishNameEdit();
                }
                else
                {
                    DehighlightSelected();
                }
            }

            myShapeFlowLayout.Focus();
        }

        private void FocusSelected()
        {
            myShapeFlowLayout.ScrollControlIntoView(_selectedThumbnail[0]);
            _selectedThumbnail[0].Highlight();
        }

        private void InitializeContextMenu()
        {
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
        }

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
                    // if user trys to import shapes but the file contains multiple categories,
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
            PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
            PowerPointPresentation pres = PowerPointPresentation.Current;

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
            PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
            PowerPointPresentation pres = PowerPointPresentation.Current;

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

            PaneReload(true);
            loadingDialog.Close();

            return true;
        }

        private void MultiSelectClickHandler(LabeledThumbnail clickedThumbnail)
        {
            if (MouseButtons != MouseButtons.Left &&
                MouseButtons != MouseButtons.Right)
            {
                return;
            }

            // for right click, if selection > 1, the context menu should appear with selection
            // remained, else we should change the focus. Specially, when selection > 1, some of
            // the options in the context menu serves for the clicked item, such as rename.
            if (MouseButtons == MouseButtons.Right)
            {
                if (_selectedThumbnail.Count > 1 &&
                    _selectedThumbnail.Contains(clickedThumbnail))
                {
                    _selectedThumbnail.Remove(clickedThumbnail);
                    _selectedThumbnail.Insert(0, clickedThumbnail);

                    return;
                }
            }

            // if Ctrl key is not holding, i.e. not doing multi-selecting, all highlighed
            // thumbnail should be dehighlighted
            if (!ModifierKeys.HasFlag(Keys.Control))
            {
                foreach (LabeledThumbnail thumbnail in _selectedThumbnail)
                {
                    thumbnail.DeHighlight();
                }

                _selectedThumbnail.Clear();
            }

            if (!_selectedThumbnail.Contains(clickedThumbnail))
            {
                // highlight the thumbnail and add the clicked thumbnail to the collection
                clickedThumbnail.Highlight();

                _selectedThumbnail.Insert(0, clickedThumbnail);
                FocusSelected();
            }
            else
            {
                // turn off the highlighting if the clicked thumbnail is currently highlighted
                if (ModifierKeys.HasFlag(Keys.Control))
                {
                    clickedThumbnail.DeHighlight();

                    _clickOnSelected = false;
                    _selectedThumbnail.Remove(clickedThumbnail);
                }
            }
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

            IOrderedEnumerable<string> shapes = Directory.EnumerateFiles(CurrentShapeFolderPath, "*.png").OrderBy(item => item, _stringComparer);

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

            if (myShapeFlowLayout.Controls.Count == 0)
            {
                ShowNoShapeMessage();
            }

            DehighlightSelected();
        }

        private void RegulateSelectionRectPoint(ref Point p)
        {
            if (p.X < 0)
            {
                p.X = 0;
            }
            else
                if (p.X > myShapeFlowLayout.Width)
            {
                p.X = myShapeFlowLayout.Width;
            }

            if (p.Y < 0)
            {
                p.Y = 0;
            }
            else
                if (p.Y > myShapeFlowLayout.Height)
            {
                p.Y = myShapeFlowLayout.Height;
            }
        }

        private void RemoveThumbnail(LabeledThumbnail thumbnail, bool removeSelection = true)
        {
            if (removeSelection &&
                _selectedThumbnail.Contains(thumbnail))
            {
                _selectedThumbnail.Remove(thumbnail);
            }

            myShapeFlowLayout.Controls.Remove(thumbnail);

            if (myShapeFlowLayout.Controls.Count == 0)
            {
                ShowNoShapeMessage();
            }
        }

        private void RenameThumbnail(string oldName, LabeledThumbnail labeledThumbnail)
        {
            if (oldName == labeledThumbnail.NameLabel)
            {
                return;
            }

            string newPath = labeledThumbnail.ImagePath.Replace(@"\" + oldName, @"\" + labeledThumbnail.NameLabel);

            File.Move(labeledThumbnail.ImagePath, newPath);
            labeledThumbnail.ImagePath = newPath;

            Globals.ThisAddIn.ShapePresentation.RenameShape(oldName, labeledThumbnail.NameLabel);

            Globals.ThisAddIn.SyncShapeRename(oldName, labeledThumbnail.NameLabel, CurrentCategory);
        }

        private void ReorderThumbnail(LabeledThumbnail labeledThumbnail)
        {
            int index = FindLabeledThumbnailIndex(labeledThumbnail.NameLabel);

            // if the current control is the only control or something goes wrong, don't need
            // to reorder
            if (index == -1 ||
                index >= myShapeFlowLayout.Controls.Count)
            {
                return;
            }

            myShapeFlowLayout.Controls.SetChildIndex(labeledThumbnail, index);
        }

        private void ShowNoShapeMessage()
        {
            if (_noShapePanel.Controls.Count == 0)
            {
                _noShapePanel.Controls.Add(_noShapeLabelFirstLine);
                _noShapePanel.Controls.Add(_noShapeLabelSecondLine);
            }

            myShapeFlowLayout.Controls.Add(_noShapePanel);
        }
        # endregion

        # region Event Handlers
        private void CategoryBoxSelectedIndexChanged(object sender, EventArgs e)
        {
            int selectedIndex = categoryBox.SelectedIndex;
            string selectedCategory = _categoryBinding[selectedIndex].ToString();

            CurrentCategory = selectedCategory;
            Globals.ThisAddIn.ShapePresentation.DefaultCategory = selectedCategory;
            PaneReload(true);
        }

        private void CategoryBoxOwnerDraw(object sender, DrawItemEventArgs e)
        {
            ComboBox comboBox = sender as ComboBox;

            if (comboBox == null ||
                e.Index == -1)
            {
                return;
            }

            Font font = comboBox.Font;
            string text = (string)_categoryBinding[e.Index];

            if (text == Globals.ThisAddIn.ShapesLabConfig.DefaultCategory)
            {
                text += " (default)";
                font = new Font(font, FontStyle.Bold);
            }

            using (SolidBrush brush = new SolidBrush(e.ForeColor))
            {
                e.DrawBackground();
                e.Graphics.DrawString(text, font, brush, e.Bounds);
            }

            int desiredWidth = Width - label1.Width - 60;
            comboBox.Width = desiredWidth > 0 ? desiredWidth : 0;
        }

        private void CopyContextMenuStripLeaveEvent(object sender, EventArgs e)
        {
            copyToToolStripMenuItem.HideDropDown();
        }

        private void CopyContextMenuStripOnEvent(object sender, EventArgs e)
        {
            copyToToolStripMenuItem.DropDownItems.Clear();

            foreach (string category in _categoryBinding.List)
            {
                if (category != CurrentCategory)
                {
                    ToolStripItem item = copyToToolStripMenuItem.DropDownItems.Add(category);
                    item.Click += CopyContextMenuStripSubMenuClick;
                }
            }

            copyToToolStripMenuItem.ShowDropDown();
        }

        private void CopyContextMenuStripSubMenuClick(object sender, EventArgs e)
        {
            ToolStripItem item = sender as ToolStripItem;

            if (item == null)
            {
                return;
            }

            string categoryName = item.Text;

            GraphicsUtil.SuspendDrawing(myShapeFlowLayout);

            foreach (LabeledThumbnail thumbnail in _selectedThumbnail)
            {
                string shapeName = thumbnail.NameLabel;

                string oriPath = Path.Combine(CurrentShapeFolderPath, shapeName) + ".png";
                string destPath = Path.Combine(ShapeRootFolderPath, categoryName, shapeName) + ".png";

                // if we have an identical name in the destination category, we won't allow
                // moving
                if (File.Exists(destPath))
                {
                    MessageBox.Show(string.Format(TextCollection.ShapesLabText.ErrorSameShapeNameInDestination,
                                    shapeName,
                                    categoryName));
                    break;
                }

                // Utilises deprecated classes as CustomShapePane does not utilise ActionFramework
                PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
                PowerPointPresentation pres = PowerPointPresentation.Current;

                // move shape in ShapeGallery to correct place
                Globals.ThisAddIn.ShapePresentation.CopyShapeToCategory(pres, currentSlide, shapeName, categoryName);

                // move shape on the disk to correct place
                File.Copy(oriPath, destPath);

                Globals.ThisAddIn.SyncShapeAdd(shapeName, destPath, categoryName);
            }

            GraphicsUtil.ResumeDrawing(myShapeFlowLayout);
            _selectedThumbnail.Clear();
        }

        private void CustomShapePaneClick(object sender, EventArgs e)
        {
            FlowlayoutClick();
        }

        private void FlowlayoutContextMenuStripItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            ToolStripItem item = e.ClickedItem;

            if (item.Name.Contains("settings"))
            {
                ContextMenuStripSettingsClicked();
            }
            else if (item.Name.Contains("addCategory"))
            {
                ContextMenuStripAddCategoryClicked();
            }
            else if (item.Name.Contains("removeCategory"))
            {
                ContextMenuStripRemoveCategoryClicked();
            }
            else if (item.Name.Contains("renameCategory"))
            {
                ContextMenuStripRenameCategoryClicked();
            }
            else if (item.Name.Contains("setAsDefault"))
            {
                ContextMenuStripSetAsDefaultCategoryClicked();
            }
            else if (item.Name.Contains("importCategory"))
            {
                ContextMenuStripImportCategoryClicked();
            }
            else if (item.Name.Contains("importShape"))
            {
                ContextMenuStripImportShapesClicked();
            }
        }

        private void FlowLayoutMouseDownHandler(object sender, MouseEventArgs e)
        {
            if (!ModifierKeys.HasFlag(Keys.Control))
            {
                FlowlayoutClick();
            }

            _isPanelMouseDown = true;
            _isPanelDrawingFinish = false;
            _startPosition = e.Location;

            _selectRect.Location = myShapeFlowLayout.PointToScreen(e.Location);
            _selectRect.Size = new Size(0, 0);
            _selectRect.BringToFront();
            _selectRect.Show();

            _selectingThumbnail.Clear();
        }

        private void FlowLayoutMouseEnterHandler(object sender, EventArgs e)
        {
            if (_selectedThumbnail != null &&
                _selectedThumbnail.Count != 0 &&
                _selectedThumbnail[0].State != LabeledThumbnail.Status.Editing)
            {
                myShapeFlowLayout.Focus();
            }
        }

        private void FlowLayoutMouseMoveHandler(object sender, MouseEventArgs e)
        {
            if (_isPanelMouseDown)
            {
                _curPosition = e.Location;
                Rectangle rect = CreateRect(_curPosition, _startPosition);

                _selectRect.Size = rect.Size;
                _selectRect.Location = myShapeFlowLayout.PointToScreen(rect.Location);

                foreach (Control control in myShapeFlowLayout.Controls)
                {
                    if (!(control is LabeledThumbnail))
                    {
                        continue;
                    }

                    LabeledThumbnail labeledThumbnail = control as LabeledThumbnail;
                    Rectangle labeledThumbnailRect =
                        myShapeFlowLayout.RectangleToClient(
                            labeledThumbnail.RectangleToScreen(labeledThumbnail.ClientRectangle));

                    if (labeledThumbnailRect.IntersectsWith(rect))
                    {
                        if (!_selectingThumbnail.Contains(labeledThumbnail))
                        {
                            labeledThumbnail.ToggleHighlight();
                            _selectingThumbnail.Add(labeledThumbnail);
                        }
                    }
                    else
                    {
                        if (_selectingThumbnail.Contains(labeledThumbnail))
                        {
                            labeledThumbnail.ToggleHighlight();
                            _selectingThumbnail.Remove(labeledThumbnail);
                        }
                    }
                }

                myShapeFlowLayout.Invalidate();
            }
        }

        private void FlowLayoutMouseUpHandler(object sender, MouseEventArgs e)
        {
            _isPanelMouseDown = false;
            _isPanelDrawingFinish = true;
            _selectRect.Hide();

            foreach (LabeledThumbnail thumbnail in _selectingThumbnail)
            {
                if (_selectedThumbnail.Contains(thumbnail))
                {
                    _selectedThumbnail.Remove(thumbnail);
                }
                else
                {
                    _selectedThumbnail.Add(thumbnail);
                }
            }

            if (_selectedThumbnail.Count != 0)
            {
                _selectedThumbnail = _selectedThumbnail.OrderBy(item => item.NameLabel, _stringComparer).ToList();
            }
        }

        private void FlowLayoutPaintHandler(object sender, PaintEventArgs e)
        {
            if (_isPanelMouseDown)
            {
                Rectangle rect = CreateRect(_curPosition, _startPosition);

                using (SolidBrush brush = new SolidBrush(Color.FromArgb(100, 0, 0, 255)))
                {
                    e.Graphics.FillRectangle(brush, rect);
                }

                using (Pen pen = new Pen(Color.FromArgb(200, 0, 0, 255)))
                {
                    e.Graphics.DrawRectangle(pen, rect);
                }
            }

            if (_isPanelDrawingFinish)
            {
                e.Graphics.Clear(myShapeFlowLayout.BackColor);
            }
        }

        private void LabeledThumbnailClick(object sender, MouseEventArgs e)
        {
            if (sender == null || !(sender is LabeledThumbnail))
            {
                MessageBox.Show(ShapesLabText.ErrorNoPanelSelected);
                return;
            }

            _clicks++;

            if (flowlayoutContextMenuStrip.Visible)
            {
                flowlayoutContextMenuStrip.Hide();
            }

            // only first click will be entertained
            if (!_firstClick)
            {
                return;
            }

            myShapeFlowLayout.Focus();

            _firstClick = false;
            _isLeftButton = e.Button == MouseButtons.Left;

            FirstClickOnThumbnail(sender as LabeledThumbnail);

            // if it's left button click, we need to wait for potential second click
            _timer.Start();
        }

        private void LabeledThumbnailDoubleClick(object sender, EventArgs e)
        {
            if (sender == null || !(sender is LabeledThumbnail))
            {
                MessageBox.Show(ShapesLabText.ErrorNoPanelSelected);
                return;
            }

            LabeledThumbnail clickedThumbnail = sender as LabeledThumbnail;

            string shapeName = clickedThumbnail.NameLabel;
            PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;

            // Utilises deprecated PowerPointPresentation class as CustomShapePane does not utilise ActionFramework
            PowerPointPresentation pres = PowerPointPresentation.Current;

            if (currentSlide != null)
            {
                Globals.ThisAddIn.Application.StartNewUndoEntry();

                ClipboardUtil.RestoreClipboardAfterAction(() =>
                {
                    Globals.ThisAddIn.ShapePresentation.CopyShape(shapeName);
                    currentSlide.Shapes.Paste().Select();
                    return ClipboardUtil.ClipboardRestoreSuccess;
                }, pres, currentSlide);
            }
            else
            {
                MessageBox.Show(ShapesLabText.ErrorViewTypeNotSupported);
            }
        }

        private void NameEditFinishHandler(object sender, string oldName)
        {
            LabeledThumbnail labeledThumbnail = sender as LabeledThumbnail;

            // by right, name change only happens when the labeled thumbnail is selected.
            // Therfore, if the notifier doesn't come from the selected object, something
            // goes wrong.
            if (labeledThumbnail == null ||
                (_selectedThumbnail.Count != 0 &&
                labeledThumbnail != _selectedThumbnail[0]))
            {
                return;
            }

            // if name changed, rename the shape in shape gallery and the file on disk
            RenameThumbnail(oldName, labeledThumbnail);

            // put the labeled thumbnail to correct position
            ReorderThumbnail(labeledThumbnail);

            // select the thumbnail and scroll into view
            FocusSelected();
        }

        private void MoveContextMenuStripLeaveEvent(object sender, EventArgs e)
        {
            moveShapeToolStripMenuItem.HideDropDown();
        }

        private void MoveContextMenuStripOnEvent(object sender, EventArgs e)
        {
            moveShapeToolStripMenuItem.DropDownItems.Clear();

            foreach (string category in _categoryBinding.List)
            {
                if (category != CurrentCategory)
                {
                    ToolStripItem item = moveShapeToolStripMenuItem.DropDownItems.Add(category);
                    item.Click += MoveContextMenuStripSubMenuClick;
                }
            }

            moveShapeToolStripMenuItem.ShowDropDown();
        }

        private void MoveContextMenuStripSubMenuClick(object sender, EventArgs e)
        {
            ToolStripItem item = sender as ToolStripItem;

            if (item == null)
            {
                return;
            }

            string categoryName = item.Text;

            GraphicsUtil.SuspendDrawing(myShapeFlowLayout);

            foreach (LabeledThumbnail thumbnail in _selectedThumbnail)
            {
                string shapeName = thumbnail.NameLabel;

                string oriPath = Path.Combine(CurrentShapeFolderPath, shapeName) + ".png";
                string destPath = Path.Combine(ShapeRootFolderPath, categoryName, shapeName) + ".png";

                // if we have an identical name in the destination category, we won't allow
                // moving
                if (File.Exists(destPath))
                {
                    MessageBox.Show(string.Format(TextCollection.ShapesLabText.ErrorSameShapeNameInDestination, 
                                    shapeName, 
                                    categoryName));
                    break;
                }

                // Utilises deprecated classes as CustomShapePane does not utilise ActionFramework
                PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
                PowerPointPresentation pres = PowerPointPresentation.Current;

                // move shape in ShapeGallery to correct place
                Globals.ThisAddIn.ShapePresentation.MoveShapeToCategory(pres, currentSlide, shapeName, categoryName);

                // move shape on the disk to correct place
                File.Move(oriPath, destPath);

                // remove the thumbnail on the pane
                RemoveThumbnail(thumbnail, false);

                Globals.ThisAddIn.SyncShapeRemove(shapeName, CurrentCategory);
                Globals.ThisAddIn.SyncShapeAdd(shapeName, destPath, categoryName);
            }

            GraphicsUtil.ResumeDrawing(myShapeFlowLayout);
            _selectedThumbnail.Clear();
        }

        private void ThumbnailContextMenuStripItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            ToolStripItem item = e.ClickedItem;

            if (item.Name.Contains("remove"))
            {
                ContextMenuStripRemoveClicked();
            }
            else
            if (item.Name.Contains("edit"))
            {
                ContextMenuStripEditClicked();
            }
            else
            if (item.Name.Contains("add"))
            {
                ContextMenuStripAddClicked();
            }
        }

        private void SelectionChanged(Selection selection)
        {
            if ((selection == null) || (selection.Type == PpSelectionType.ppSelectionNone)
                || (selection.Type == PpSelectionType.ppSelectionSlides))
            {
                DisableAddShapesButton();
            }
            else
            {
                EnableAddShapesButton();
            }
        }

        private void ThumbnailContextMenuStripOpening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (Categories.Count < 2)
            {
                moveShapeToolStripMenuItem.Enabled = false;
                copyToToolStripMenuItem.Enabled = false;

                moveShapeToolStripMenuItem.DropDownItems.Clear();
                copyToToolStripMenuItem.DropDownItems.Clear();
            }
            else
            {
                moveShapeToolStripMenuItem.Enabled = true;
                copyToToolStripMenuItem.Enabled = true;

                // add a dummy entry to show right arrow
                moveShapeToolStripMenuItem.DropDownItems.Add("");
                copyToToolStripMenuItem.DropDownItems.Add("");
            }
        }

        private void TimerTickHandler(object sender, EventArgs args)
        {
            _timer.Stop();

            // if we got only 1 click in a threshold value, we take it as a single click
            if (_clicks == 1 &&
                _isLeftButton &&
                _clickOnSelected)
            {
                _selectedThumbnail[0].StartNameEdit();
            }

            ClickTimerReset();
        }
        #endregion

        #region search box appearance and behaviors
        /*
        private bool _searchBoxFocused = false;
        protected override void OnLoad(EventArgs e)
        {
            var searchButton = new Button();

            searchButton.Size = new Size(25, searchBox.ClientSize.Height + 2);
            searchButton.Location = new Point(searchBox.ClientSize.Width - searchButton.Width, -1);
            searchButton.Image = Properties.Resources.EditNameContext;
            searchButton.Cursor = Cursors.Hand;

            searchBox.Controls.Add(searchButton);

            // send EM_SETMARGINS to text box to prevent words from going under the button
            Native.SendMessage(searchBox.Handle, 0xd3, (IntPtr)2, (IntPtr)(searchButton.Width << 16));
            base.OnLoad(e);
        }

        private void SearchBoxLeave(object sender, EventArgs e)
        {
            _searchBoxFocused = false;
        }

        private void SearchBoxEnter(object sender, EventArgs e)
        {
            // only when user mouse down & up in the text box we do highlighting
            if (MouseButtons == MouseButtons.None)
            {
                searchBox.SelectAll();
                _searchBoxFocused = true;
            }
        }

        private void SearchBoxMouseUp(object sender, MouseEventArgs e)
        {
            if (!_searchBoxFocused)
            {
                if (searchBox.SelectionLength == 0)
                {
                    searchBox.SelectAll();
                }

                _searchBoxFocused = true;
            }
        }
        */
        #endregion

        #region GUI Handles

        private void AddShapeButton_Click(object sender, EventArgs e)
        {
            Selection selection = ActionFrameworkExtensions.GetCurrentSelection();
            ThisAddIn addIn = ActionFrameworkExtensions.GetAddIn();

            AddShapeFromSelection(selection, addIn);
        }

        // A disabled button cannot respond to any events.
        // Thus we register the event to the pane and when the mouse moves over
        // the button, the tool tip will display.
        private void CustomShapePane_MouseMove(object sender, MouseEventArgs e)
        {
            Control parent = sender as Control;
            if (parent == null)
            {
                return;
            }
            Control ctrl = parent.GetChildAtPoint(e.Location);
            if (ctrl != null)
            {
                if (ctrl.Visible && toolTip1.Tag == null)
                {
                    if (!_toolTipShown)
                    {
                        toolTip1.Show(ShapesLabText.DisabledAddShapeToolTip, ctrl, ctrl.Width / 2, ctrl.Height / 2);
                        toolTip1.Tag = ctrl;
                        _toolTipShown = true;
                    }
                }
            }
            else
            {
                Control toolTipCtrl = toolTip1.Tag as Control;
                if (toolTipCtrl != null)
                {
                    toolTip1.Hide(toolTipCtrl);
                    toolTip1.Tag = null;
                    _toolTipShown = false;
                }
            }
        }
        #endregion

        #region ToolTip
        private void InitToolTipControl()
        {
            toolTip1.SetToolTip(addShapeButton, ShapesLabText.AddShapeToolTip);
        }
        #endregion
    }
}
