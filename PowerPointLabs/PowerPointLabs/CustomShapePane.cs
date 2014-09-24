using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using PPExtraEventHelper;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;
using Graphics = PowerPointLabs.Utils.Graphics;

namespace PowerPointLabs
{
    public partial class CustomShapePane : UserControl
    {
        private const string DefaultShapeNameFormat = @"My Shape Untitled {0}";
        private const string DefaultShapeNameSearchRegex = @"^My Shape Untitled (\d+)$";
        private const string ShapeFileDialogFilter =
            "PowerPointLabs Shapes File|*.pptlabsshapes;*.pptx";

        private readonly int _doubleClickTimeSpan = SystemInformation.DoubleClickTime;
        private int _clicks;

        private bool _firstTimeLoading = true;
        private bool _firstClick = true;
        private bool _clickOnSelected;
        private bool _isLeftButton;

        private bool _isPanelMouseDown;
        private bool _isPanelDrawingFinish;
        private Point _startPosition;
        private Point _curPosition;

        private readonly SelectionRectangle _selectRect = new SelectionRectangle();

        private readonly BindingSource _categoryBinding;

        private List<LabeledThumbnail> _selectedThumbnail;

        private readonly Timer _timer;

        private readonly Comparers.AtomicNumberStringCompare _stringComparer = new Comparers.AtomicNumberStringCompare();

        # region Properties
        public string NextDefaultFullName
        {
            get { return CurrentShapeFolderPath + @"\" +
                         NextDefaultNameWithoutExtension + ".png"; }
        }

        public string NextDefaultNameWithoutExtension
        {
            get
            {
                var labelNames =
                    myShapeFlowLayout.Controls.OfType<LabeledThumbnail>().Select(control => control.NameLable).ToList();

                var nextNum = Common.NextDefaultNumber(labelNames, new Regex(DefaultShapeNameSearchRegex));

                return string.Format(DefaultShapeNameFormat, nextNum);
            }
        }

        public List<string> Categories { get; private set; }

        public string CurrentCategory { get; set; }

        public string CurrentShapeFullName
        {
            get { return CurrentShapeFolderPath + @"\" +
                         CurrentShapeNameWithoutExtension + ".png"; }
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

                return _selectedThumbnail[0].NameLable;
            }
        }

        public List<string> Shapes
        {
            get
            {
                var shapeList = new List<string>();

                if (myShapeFlowLayout.Controls.Count == 0 ||
                    myShapeFlowLayout.Controls.Contains(_noShapePanel))
                {
                    return shapeList;
                }

                shapeList.AddRange(from LabeledThumbnail labelThumbnail in myShapeFlowLayout.Controls
                                   select labelThumbnail.NameLable);

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
                var createParams = base.CreateParams;

                // do this optimization only for office 2010 since painting speed on 2013 is
                // really slow
                if (Globals.ThisAddIn.Application.Version == Globals.ThisAddIn.OfficeVersion2010)
                {
                    createParams.ExStyle |= (int)Native.Message.WS_EX_COMPOSITED;  // Turn on WS_EX_COMPOSITED
                }

                return createParams;
            }
        }
        # endregion

        # region Constructors
        public CustomShapePane(string shapeRootFolderPath, string defaultShapeCategoryName)
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.DoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
            InitializeComponent();

            InitializeContextMenu();

            _selectedThumbnail = new List<LabeledThumbnail>();

            ShapeRootFolderPath = shapeRootFolderPath;

            CurrentCategory = defaultShapeCategoryName;
            Categories = new List<string>(Globals.ThisAddIn.ShapePresentation.Categories);
            _categoryBinding = new BindingSource { DataSource = Categories };
            categoryBox.DataSource = _categoryBinding;

            for (var i = 0; i < Categories.Count; i ++)
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

            //myShapeFlowLayout.Paint += FlowLayoutPaintHandler;
        }
        # endregion

        # region API
        public void AddCustomShape(string shapeName, string shapePath, bool immediateEditing)
        {
            DehighlightSelected();

            var labeledThumbnail = new LabeledThumbnail(shapePath, shapeName) { ContextMenuStrip = shapeContextMenuStrip };

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
            var labeledThumbnail = FindLabeledThumbnail(shapeName);

            if (labeledThumbnail == null)
            {
                return;
            }

            // remove shape from task pane
            RemoveThumbnail(labeledThumbnail);
        }

        public void RenameCustomShape(string oldShapeName, string newShapeName)
        {
            var labeledThumbnail = FindLabeledThumbnail(oldShapeName);

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
            if (Globals.ThisAddIn.Application.Version == Globals.ThisAddIn.OfficeVersion2013)
            {
                Graphics.SuspendDrawing(myShapeFlowLayout);
            }
            
            // emptize the panel and load shapes from folder
            myShapeFlowLayout.Controls.Clear();
            PrepareShapes();
            
            // scroll the view to show the first item, and focus the flowlayout to enable
            // scroll if applicable
            myShapeFlowLayout.ScrollControlIntoView(myShapeFlowLayout.Controls[0]);
            myShapeFlowLayout.Focus();

            // double buffer ends
            if (Globals.ThisAddIn.Application.Version == Globals.ThisAddIn.OfficeVersion2013)
            {
                Graphics.ResumeDrawing(myShapeFlowLayout);
            }

            _firstTimeLoading = false;
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

        private void ContextMenuStripAddCategoryClicked()
        {
            var categoryInfoDialog = new ShapesLabCategoryInfoForm(string.Empty);

            categoryInfoDialog.ShowDialog();

            if (categoryInfoDialog.UserOption == ShapesLabCategoryInfoForm.Option.Ok)
            {
                var categoryName = categoryInfoDialog.CategoryName;

                Globals.ThisAddIn.ShapePresentation.AddCategory(categoryName);

                _categoryBinding.Add(categoryName);

                categoryBox.SelectedIndex = _categoryBinding.Count - 1;
            }

            myShapeFlowLayout.Focus();
        }

        private void ContextMenuStripEditClicked()
        {
            if (_selectedThumbnail == null)
            {
                MessageBox.Show(TextCollection.CustomShapeNoPanelSelectedError);
                return;
            }

            _selectedThumbnail[0].StartNameEdit();
        }

        private void ContextMenuStripImportCategoryClicked()
        {
            var fileDialog = new OpenFileDialog
                                 {
                                     Filter = ShapeFileDialogFilter,
                                     Multiselect = false
                                 };
            
            flowlayoutContextMenuStrip.Hide();

            if (fileDialog.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }

            var importFilePath = fileDialog.FileName;
            var importFileName = new FileInfo(importFilePath).Name;
            var importFileNameNoExtension = importFileName.Substring(0, importFileName.LastIndexOf('.'));
            var importFileCopyPath = Path.Combine(ShapeRootFolderPath, importFileName);
            var sameFolder = true;

            // copy the file to the current shape root if the file is not under root 
            if (!File.Exists(importFileCopyPath))
            {
                File.Copy(importFilePath, importFileCopyPath);
                sameFolder = false;
            }

            // open the file as an imported file
            var importShapeGallery = new PowerPointShapeGalleryPresentation(ShapeRootFolderPath,
                                                                            importFileNameNoExtension)
                                         {IsImportedFile = true};
            
            if (!importShapeGallery.Open(withWindow: false, focus: false) &&
                !importShapeGallery.Opened)
            {
                MessageBox.Show(TextCollection.CustomShapeImportFileError);
            }
            else
            {
                // copy all shapes in the import shape gallery to current shape gallery
                foreach (var importCategory in importShapeGallery.Categories)
                {
                    importShapeGallery.RetrieveCategory(importCategory);
                    Globals.ThisAddIn.ShapePresentation.AppendCategoryFromClipBoard();
                    _categoryBinding.Add(importCategory);
                }
            }

            importShapeGallery.Close();

            // delete the import file copy
            if (!sameFolder)
            {
                if (importFileCopyPath.EndsWith(".pptx"))
                {
                    importFileCopyPath = importFileCopyPath.Replace(".pptx", ".pptlabsshapes");
                }

                FileDir.DeleteFile(importFileCopyPath);
            }

            MessageBox.Show(TextCollection.CustomShapeImportSuccess);
        }

        private void ContextMenuStripRemoveCategoryClicked()
        {
            // remove the last category will not be entertained
            if (_categoryBinding.Count == 1)
            {
                MessageBox.Show(TextCollection.CustomShapeRemoveLastCategoryError);
                return;
            }

            var categoryIndex = categoryBox.SelectedIndex;
            var categoryName = _categoryBinding[categoryIndex].ToString();
            var categoryPath = Path.Combine(ShapeRootFolderPath, categoryName);
            var isDefaultCategory = Globals.ThisAddIn.ShapesLabConfigs.DefaultCategory == CurrentCategory;

            if (isDefaultCategory)
            {
                var result =
                    MessageBox.Show(TextCollection.CustomShapeRemoveDefaultCategoryMessage,
                                    TextCollection.CustomShapeRemoveDefaultCategoryCaption,
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
                Globals.ThisAddIn.ShapesLabConfigs.DefaultCategory = (string)_categoryBinding[0];
            }
        }

        private void ContextMenuStripRemoveClicked()
        {
            if (_selectedThumbnail == null ||
                _selectedThumbnail.Count == 0)
            {
                MessageBox.Show(TextCollection.CustomShapeNoPanelSelectedError);
                return;
            }

            Graphics.SuspendDrawing(myShapeFlowLayout);

            while (_selectedThumbnail.Count > 0)
            {
                var thumbnail = _selectedThumbnail[0];
                var removedShapename = thumbnail.NameLable;

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

            Graphics.ResumeDrawing(myShapeFlowLayout);
        }

        private void ContextMenuStripRenameCategoryClicked()
        {
            var categoryInfoDialog = new ShapesLabCategoryInfoForm(CurrentCategory);

            categoryInfoDialog.ShowDialog();

            if (categoryInfoDialog.UserOption == ShapesLabCategoryInfoForm.Option.Ok)
            {
                var categoryName = categoryInfoDialog.CategoryName;

                // if current category is the default category, change ShapeConfig
                if (Globals.ThisAddIn.ShapesLabConfigs.DefaultCategory == CurrentCategory)
                {
                    Globals.ThisAddIn.ShapesLabConfigs.DefaultCategory = categoryName;
                }

                // rename the category in ShapeGallery
                Globals.ThisAddIn.ShapePresentation.RenameCategory(categoryName);
                
                // rename the category on the disk
                var newPath = Path.Combine(ShapeRootFolderPath, categoryName);
                
                try
                {
                    Directory.Move(CurrentShapeFolderPath, newPath);
                } catch (Exception)
                {
                    // this may occur when the newCategoryName.tolower() == oldCategoryName.tolower()
                }

                // rename the category in combo box
                var categoryIndex = categoryBox.SelectedIndex;
                _categoryBinding[categoryIndex] = categoryName;

                // update current category reference
                CurrentCategory = categoryName;
            }

            myShapeFlowLayout.Focus();
        }

        private void ContextMenuStripSetAsDefaultCategoryClicked()
        {
            Globals.ThisAddIn.ShapesLabConfigs.DefaultCategory = CurrentCategory;

            categoryBox.Refresh();
            flowlayoutContextMenuStrip.Hide();

            MessageBox.Show(string.Format(TextCollection.CustomeShapeSetAsDefaultCategorySuccessFormat, CurrentCategory));
        }

        private void ContextMenuStripSettingsClicked()
        {
            var settingDialog = new ShapesLabSetting(ShapeRootFolderPath);

            settingDialog.ShowDialog();

            if (settingDialog.UserOption == ShapesLabSetting.Option.Ok)
            {
                var newPath = settingDialog.DefaultSavingPath;

                if (!MigrateShapeFolder(ShapeRootFolderPath, newPath))
                {
                    return;
                }

                Globals.ThisAddIn.ShapesLabConfigs.ShapeRootFolder = newPath;

                MessageBox.Show(
                    string.Format(TextCollection.CustomeShapeSaveLocationChangedSuccessFormat, newPath),
                    TextCollection.CustomShapeSaveLocationChangedSuccessTitle, MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        private Rectangle CreateRect(Point loc1, Point loc2)
        {
            RegulateSelectionRectPoint(ref loc1);
            RegulateSelectionRectPoint(ref loc2);

            var size = new Size(Math.Abs(loc2.X - loc1.X), Math.Abs(loc2.Y - loc1.Y));
            var rect = new Rectangle(new Point(Math.Min(loc1.X, loc2.X), Math.Min(loc1.Y, loc2.Y)), size);

            return rect;
        }

        private void DehighlightSelected()
        {
            if (_selectedThumbnail == null ||
                _selectedThumbnail.Count == 0)
            {
                return;
            }
            
            foreach (var thumbnail in _selectedThumbnail)
            {
                thumbnail.DeHighlight();
            }

            _selectedThumbnail.Clear();
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
                    labeledThumbnail => labeledThumbnail.NameLable == name);
        }

        private int FindLabeledThumbnailIndex(string name)
        {
            if (myShapeFlowLayout.Controls.Contains(_noShapePanel))
            {
                return -1;
            }

            var totalControl = myShapeFlowLayout.Controls.Count;
            var thisControlPosition = -1;

            for (var i = 0; i < totalControl; i ++)
            {
                var control = myShapeFlowLayout.Controls[i] as LabeledThumbnail;

                if (control == null) continue;

                // skip itself
                if (control.NameLable == name)
                {
                    thisControlPosition = i;
                    continue;
                }
                
                if (_stringComparer.Compare(control.NameLable, name) > 0)
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
            // here we have 2 cases when multiple selection is enabled:
            //
            // Left Click:
            // finish editing if we have, dehighlight all selected labels and highlight
            // clicked label;
            //
            // Right Click:
            // finish editing if we have, keep all highlight and set the clicked label
            // as default.

            // common part, end editing
            if (_selectedThumbnail != null)
            {
                if (_selectedThumbnail.Count != 0)
                {
                    if (_selectedThumbnail[0].State == LabeledThumbnail.Status.Editing)
                    {
                        _selectedThumbnail[0].FinishNameEdit();
                    }
                    else
                        if (_selectedThumbnail[0] == clickedThumbnail)
                        {
                            _clickOnSelected = true;
                        }

                    if (!_selectedThumbnail.Contains(clickedThumbnail) ||
                        MouseButtons == MouseButtons.Left)
                    {
                        foreach (var thumbnail in _selectedThumbnail)
                        {
                            thumbnail.DeHighlight();
                        }

                        _selectedThumbnail.Clear();

                        clickedThumbnail.Highlight();
                    }
                    else
                        if (MouseButtons == MouseButtons.Right)
                        {
                            _selectedThumbnail.Remove(clickedThumbnail);
                        }
                }

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
            addToSlideToolStripMenuItem.Text = TextCollection.CustomShapeShapeContextStripAddToSlide;
            editNameToolStripMenuItem.Text = TextCollection.CustomShapeShapeContextStripEditName;
            moveShapeToolStripMenuItem.Text = TextCollection.CustomShapeShapeContextStripMoveShape;
            removeShapeToolStripMenuItem.Text = TextCollection.CustomShapeShapeContextStripRemoveShape;
            copyToToolStripMenuItem.Text = TextCollection.CustomShapeShapeContextStripCopyShape;

            addCategoryToolStripMenuItem.Text = TextCollection.CustomShapeCategoryContextStripAddCategory;
            removeCategoryToolStripMenuItem.Text = TextCollection.CustomShapeCategoryContextStripRemoveCategory;
            renameCategoryToolStripMenuItem.Text = TextCollection.CustomShapeCategoryContextStripRenameCategory;
            setAsDefaultToolStripMenuItem.Text = TextCollection.CustomShapeCategoryContextStripSetAsDefaultCategory;
            settingsToolStripMenuItem.Text = TextCollection.CustomShapeCategoryContextStripCategorySettings;
            importCategoryToolStripMenuItem.Text = TextCollection.CustomShapeCategoryContextStripImportCategory;
            
            // add a dummy entry to show right arrow
            moveShapeToolStripMenuItem.DropDownItems.Add("");
            copyToToolStripMenuItem.DropDownItems.Add("");

            foreach (ToolStripMenuItem contextMenu in shapeContextMenuStrip.Items)
            {
                if (contextMenu.Text != TextCollection.CustomShapeShapeContextStripMoveShape)
                {
                    contextMenu.MouseEnter += MoveContextMenuStripLeaveEvent;
                }
                
                if (contextMenu.Text != TextCollection.CustomShapeShapeContextStripCopyShape)
                {
                    contextMenu.MouseEnter += CopyContextMenuStripLeaveEvent;
                }
            }
        }

        private bool MigrateShapeFolder(string oldPath, string newPath)
        {
            var loadingDialog = new LoadingDialog(TextCollection.CustomShapeMigratingDialogTitle,
                                                  TextCollection.CustomShapeMigratingDialogContent);
            loadingDialog.Show();
            loadingDialog.Refresh();

            // close the opening presentation
            if (Globals.ThisAddIn.ShapePresentation.Opened)
            {
                Globals.ThisAddIn.ShapePresentation.Close();
            }

            // migration only cares about if the folder has been copied to the new location entirely.
            if (!FileDir.CopyFolder(oldPath, newPath))
            {
                loadingDialog.Dispose();

                MessageBox.Show(TextCollection.CustomShapeMigrationError);

                return false;
            }

            // now we will try our best to delete the original folder, but this is not guaranteed
            // because some of the using files, such as some opening shapes, and the evil thumb.db
            if (!FileDir.DeleteFolder(oldPath))
            {
                MessageBox.Show(TextCollection.CustomShapeOriginalFolderDeletionError);
            }

            ShapeRootFolderPath = newPath;

            // modify shape gallery presentation's path and name, then open it
            Globals.ThisAddIn.ShapePresentation.Path = newPath;
            Globals.ThisAddIn.ShapePresentation.Open(withWindow: false, focus: false);
            Globals.ThisAddIn.ShapePresentation.DefaultCategory = CurrentCategory;

            PaneReload(true);
            loadingDialog.Dispose();

            return true;
        }

        private void PrepareFolder()
        {
            if (!Directory.Exists(CurrentShapeFolderPath))
            {
                Directory.CreateDirectory(CurrentShapeFolderPath);
            }
        }

        private void PrepareShapes()
        {
            PrepareFolder();

            var shapes = Directory.EnumerateFiles(CurrentShapeFolderPath, "*.png").OrderBy(item => item, _stringComparer);

            foreach (var shape in shapes)
            {
                var shapeName = Path.GetFileNameWithoutExtension(shape);

                if (shapeName == null)
                {
                    MessageBox.Show(TextCollection.CustomShapeFileNameInvalid);
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
            if (oldName == labeledThumbnail.NameLable) return;

            var newPath = labeledThumbnail.ImagePath.Replace(@"\" + oldName, @"\" + labeledThumbnail.NameLable);

            File.Move(labeledThumbnail.ImagePath, newPath);
            labeledThumbnail.ImagePath = newPath;

            Globals.ThisAddIn.ShapePresentation.RenameShape(oldName, labeledThumbnail.NameLable);

            Globals.ThisAddIn.SyncShapeRename(oldName, labeledThumbnail.NameLable, CurrentCategory);
        }

        private void ReorderThumbnail(LabeledThumbnail labeledThumbnail)
        {
            var index = FindLabeledThumbnailIndex(labeledThumbnail.NameLable);

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
            var selectedIndex = categoryBox.SelectedIndex;
            var selectedCategory = _categoryBinding[selectedIndex].ToString();

            CurrentCategory = selectedCategory;
            Globals.ThisAddIn.ShapePresentation.DefaultCategory = selectedCategory;
            PaneReload(true);
        }

        private void CategoryBoxOwnerDraw(object sender, DrawItemEventArgs e)
        {
            var comboBox = sender as ComboBox;

            if (comboBox == null ||
                e.Index == -1) return;

            var font = comboBox.Font;
            var text = (string)_categoryBinding[e.Index];

            if (text == Globals.ThisAddIn.ShapesLabConfigs.DefaultCategory)
            {
                text += " (default)";
                font = new Font(font, FontStyle.Bold);
            }

            using (var brush = new SolidBrush(e.ForeColor))
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
            if (copyToToolStripMenuItem.Tag != null &&
                (string)copyToToolStripMenuItem.Tag == CurrentCategory)
            {
                copyToToolStripMenuItem.ShowDropDown();
                return;
            }

            copyToToolStripMenuItem.DropDownItems.Clear();

            foreach (string category in _categoryBinding.List)
            {
                if (category != CurrentCategory)
                {
                    var item = copyToToolStripMenuItem.DropDownItems.Add(category);
                    item.Click += CopyContextMenuStripSubMenuClick;
                }
            }

            copyToToolStripMenuItem.ShowDropDown();
        }

        private void CopyContextMenuStripSubMenuClick(object sender, EventArgs e)
        {
            var item = sender as ToolStripItem;

            if (item == null) return;

            var categoryName = item.Text;

            Graphics.SuspendDrawing(myShapeFlowLayout);

            foreach (var thumbnail in _selectedThumbnail)
            {
                var shapeName = thumbnail.NameLable;

                var oriPath = Path.Combine(CurrentShapeFolderPath, shapeName) + ".png";
                var destPath = Path.Combine(ShapeRootFolderPath, categoryName, shapeName) + ".png";

                // if we have an identical name in the destination category, we won't allow
                // moving
                if (File.Exists(destPath))
                {
                    MessageBox.Show(string.Format("{0} exists in {1}. Please rename your shape before moving.",
                                                  shapeName,
                                                  categoryName));

                    break;
                }

                // move shape in ShapeGallery to correct place
                Globals.ThisAddIn.ShapePresentation.CopyShape(shapeName, categoryName);

                // move shape on the disk to correct place
                File.Copy(oriPath, destPath);

                Globals.ThisAddIn.SyncShapeAdd(shapeName, destPath, categoryName);
            }

            Graphics.ResumeDrawing(myShapeFlowLayout);
            _selectedThumbnail.Clear();
        }

        private void CustomShapePaneClick(object sender, EventArgs e)
        {
            FlowlayoutClick();
        }

        private void FlowlayoutContextMenuStripItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            var item = e.ClickedItem;

            if (item.Name.Contains("settings"))
            {
                ContextMenuStripSettingsClicked();
            } else
            if (item.Name.Contains("addCategory"))
            {
                ContextMenuStripAddCategoryClicked();
            } else
            if (item.Name.Contains("removeCategory"))
            {
                ContextMenuStripRemoveCategoryClicked();
            } else
            if (item.Name.Contains("renameCategory"))
            {
                ContextMenuStripRenameCategoryClicked();
            } else
            if (item.Name.Contains("setAsDefault"))
            {
                ContextMenuStripSetAsDefaultCategoryClicked();
            }
            else
            if (item.Name.Contains("import"))
            {
                ContextMenuStripImportCategoryClicked();
            }
        }

        private void FlowLayoutMouseDownHandler(object sender, MouseEventArgs e)
        {
            FlowlayoutClick();

            _isPanelMouseDown = true;
            _isPanelDrawingFinish = false;
            _startPosition = e.Location;
            _selectRect.Location = myShapeFlowLayout.PointToScreen(e.Location);
            _selectRect.Size = new Size(0, 0);
            _selectRect.BringToFront();
            _selectRect.Show();
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
                var rect = CreateRect(_curPosition, _startPosition);

                _selectRect.Size = rect.Size;
                _selectRect.Location = myShapeFlowLayout.PointToScreen(rect.Location);
                
                foreach (Control control in myShapeFlowLayout.Controls)
                {
                    if (!(control is LabeledThumbnail)) continue;

                    var labeledThumbnail = control as LabeledThumbnail;
                    var labeledThumbnailRect =
                        myShapeFlowLayout.RectangleToClient(
                            labeledThumbnail.RectangleToScreen(labeledThumbnail.ClientRectangle));

                    if (labeledThumbnailRect.IntersectsWith(rect))
                    {
                        if (!_selectedThumbnail.Contains(labeledThumbnail))
                        {
                            labeledThumbnail.Highlight();
                            _selectedThumbnail.Add(labeledThumbnail);
                        }
                    }
                    else
                    {
                        if (labeledThumbnail.Highlighed)
                        {
                            labeledThumbnail.DeHighlight();
                            _selectedThumbnail.Remove(labeledThumbnail);
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

            if (_selectedThumbnail.Count != 0)
            {
                _selectedThumbnail = _selectedThumbnail.OrderBy(item => item.NameLable, _stringComparer).ToList();
            }
        }

        private void FlowLayoutPaintHandler(object sender, PaintEventArgs e)
        {
            if (_isPanelMouseDown)
            {
                var rect = CreateRect(_curPosition, _startPosition);

                using (var brush = new SolidBrush(Color.FromArgb(100, 0, 0, 255)))
                {
                    e.Graphics.FillRectangle(brush, rect);
                }

                using (var pen = new Pen(Color.FromArgb(200, 0, 0, 255)))
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
                MessageBox.Show(TextCollection.CustomShapeNoPanelSelectedError);
                return;
            }

            _clicks++;

            // only first click will be entertained
            if (!_firstClick) return;

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
                MessageBox.Show(TextCollection.CustomShapeNoPanelSelectedError);
                return;
            }

            var clickedThumbnail = sender as LabeledThumbnail;

            var shapeName = clickedThumbnail.NameLable;
            var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
            
            if (currentSlide != null)
            {
                Globals.ThisAddIn.ShapePresentation.RetrieveShape(shapeName);
                currentSlide.Shapes.Paste().Select();
            }
            else
            {
                MessageBox.Show(TextCollection.CustomShapeViewTypeNotSupported);
            }
        }

        private void NameEditFinishHandler(object sender, string oldName)
        {
            var labeledThumbnail = sender as LabeledThumbnail;

            // by right, name change only happens when the labeled thumbnail is selected.
            // Therfore, if the notifier doesn't come from the selected object, something
            // goes wrong.
            if (labeledThumbnail == null ||
                (_selectedThumbnail.Count != 0 &&
                labeledThumbnail != _selectedThumbnail[0])) return;

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
            if (moveShapeToolStripMenuItem.Tag != null &&
                (string)moveShapeToolStripMenuItem.Tag == CurrentCategory)
            {
                moveShapeToolStripMenuItem.ShowDropDown();
                return;
            }

            moveShapeToolStripMenuItem.DropDownItems.Clear();

            foreach (string category in _categoryBinding.List)
            {
                if (category != CurrentCategory)
                {
                    var item = moveShapeToolStripMenuItem.DropDownItems.Add(category);
                    item.Click += MoveContextMenuStripSubMenuClick;
                }
            }

            moveShapeToolStripMenuItem.ShowDropDown();
        }

        private void MoveContextMenuStripSubMenuClick(object sender, EventArgs e)
        {
            var item = sender as ToolStripItem;

            if (item == null) return;

            var categoryName = item.Text;

            Graphics.SuspendDrawing(myShapeFlowLayout);

            foreach (var thumbnail in _selectedThumbnail)
            {
                var shapeName = thumbnail.NameLable;

                var oriPath = Path.Combine(CurrentShapeFolderPath, shapeName) + ".png";
                var destPath = Path.Combine(ShapeRootFolderPath, categoryName, shapeName) + ".png";

                // if we have an identical name in the destination category, we won't allow
                // moving
                if (File.Exists(destPath))
                {
                    MessageBox.Show(string.Format("{0} exists in {1}. Please rename your shape before moving.", shapeName,
                                                  categoryName));

                    return;
                }

                // move shape in ShapeGallery to correct place
                Globals.ThisAddIn.ShapePresentation.MoveShape(shapeName, categoryName);

                // move shape on the disk to correct place
                File.Move(oriPath, destPath);

                // remove the thumbnail on the pane
                RemoveThumbnail(thumbnail, false);

                Globals.ThisAddIn.SyncShapeRemove(shapeName, CurrentCategory);
                Globals.ThisAddIn.SyncShapeAdd(shapeName, destPath, categoryName);
            }

            Graphics.ResumeDrawing(myShapeFlowLayout);
            _selectedThumbnail.Clear();
        }

        private void ThumbnailContextMenuStripItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            var item = e.ClickedItem;

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
                LabeledThumbnailDoubleClick(_selectedThumbnail, null);
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
        # endregion

        # region search box appearance and behaviors
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
        # endregion
    }
}
