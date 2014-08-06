using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using PPExtraEventHelper;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;

namespace PowerPointLabs
{
    public partial class CustomShapePane : UserControl
    {
        private const string DefaultShapeNameFormat = @"My Shape Untitled {0}";
        private const string DefaultShapeNameSearchRegex = @"^My Shape Untitled (\d+)$";

        private readonly string _shapeRootFolderPathConfigFile = Path.Combine(Globals.ThisAddIn.AppDataFolder,
                                                                              Globals.ThisAddIn.ShapeRootFolderConfigFileName);

        private readonly int _doubleClickTimeSpan = SystemInformation.DoubleClickTime;
        
        private LabeledThumbnail _selectedThumbnail;

        private bool _firstTimeLoading = true;
        private bool _firstClick = true;
        private bool _clickOnSelected;
        private bool _isLeftButton;

        private int _clicks;

        private readonly Timer _timer;

        private readonly AtomicNumberStringCompare _stringComparer = new AtomicNumberStringCompare();

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
                var temp = 0;
                var min = int.MaxValue;
                var match = new Regex(DefaultShapeNameSearchRegex);

                foreach (Control control in myShapeFlowLayout.Controls)
                {
                    if (!(control is LabeledThumbnail)) continue;

                    var labeledThumbnail = control as LabeledThumbnail;

                    if (match.IsMatch(labeledThumbnail.NameLable))
                    {
                        var currentCnt = int.Parse(match.Match(labeledThumbnail.NameLable).Groups[1].Value);

                        if (currentCnt - temp != 1)
                        {
                            min = Math.Min(min, temp);
                        }
                        
                        temp = currentCnt;
                    }
                }

                if (min == int.MaxValue)
                {
                    return string.Format(DefaultShapeNameFormat, temp + 1);
                }

                return string.Format(DefaultShapeNameFormat, min + 1);
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
                if (_selectedThumbnail == null)
                {
                    return null;
                }

                return _selectedThumbnail.NameLable;
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
        # endregion

        # region Constructors
        public CustomShapePane(string shapeRootFolderPath, string defaultShapeCategoryName)
        {
            InitializeComponent();

            ShapeRootFolderPath = shapeRootFolderPath;

            CurrentCategory = defaultShapeCategoryName;
            Categories = new List<string> {CurrentCategory};

            _timer = new Timer { Interval = _doubleClickTimeSpan };
            _timer.Tick += TimerTickHandler;

            ShowNoShapeMessage();
            
            myShapeFlowLayout.AutoSize = true;
            myShapeFlowLayout.Click += FlowlayoutClick;
        }
        # endregion

        # region API
        public void AddCustomShape(string shapeName, string shapeFullName, bool immediateEditing)
        {
            DehighlightSelected();

            var labeledThumbnail = new LabeledThumbnail(shapeFullName, shapeName) {ContextMenuStrip = shapeContextMenuStrip};

            labeledThumbnail.Click += LabeledThumbnailClick;
            labeledThumbnail.DoubleClick += LabeledThumbnailDoubleClick;
            labeledThumbnail.NameEditFinish += NameEditFinishHandler;

            myShapeFlowLayout.Controls.Add(labeledThumbnail);

            if (myShapeFlowLayout.Controls.Contains(_noShapePanel))
            {
                myShapeFlowLayout.Controls.Remove(_noShapePanel);
            }

            myShapeFlowLayout.ScrollControlIntoView(labeledThumbnail);

            _selectedThumbnail = labeledThumbnail;

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

            // free selected thumbnail
            if (labeledThumbnail == _selectedThumbnail)
            {
                _selectedThumbnail = null;
            }

            myShapeFlowLayout.Controls.Remove(labeledThumbnail);
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
            if (labeledThumbnail == _selectedThumbnail)
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

            Native.SendMessage(myShapeFlowLayout.Handle, (uint) Native.Message.WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero);
            
            // emptize the panel and load shapes from folder
            myShapeFlowLayout.Controls.Clear();
            PrepareShapes();

            Native.SendMessage(myShapeFlowLayout.Handle, (uint) Native.Message.WM_SETREDRAW, new IntPtr(1), IntPtr.Zero);
            myShapeFlowLayout.Refresh();

            Refresh();

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

        private void ContextMenuStripRemoveClicked()
        {
            if (_selectedThumbnail == null)
            {
                MessageBox.Show(TextCollection.CustomShapeNoPanelSelectedError);
                return;
            }

            var removedShapename = _selectedThumbnail.NameLable;

            // remove shape from shape gallery
            Globals.ThisAddIn.ShapePresentation.RemoveShape(CurrentShapeNameWithoutExtension);

            // remove shape from disk and shape gallery
            File.Delete(CurrentShapeFullName);

            // remove shape from task pane
            myShapeFlowLayout.Controls.Remove(_selectedThumbnail);
            _selectedThumbnail = null;

            // sync shape removing among all task panes
            Globals.ThisAddIn.SyncShapeRemove(removedShapename);

            if (myShapeFlowLayout.Controls.Count == 0)
            {
                ShowNoShapeMessage();
            }
        }

        private void ContextMenuStripEditClicked()
        {
            if (_selectedThumbnail == null)
            {
                MessageBox.Show(TextCollection.CustomShapeNoPanelSelectedError);
                return;
            }

            _selectedThumbnail.StartNameEdit();
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

                ShapeRootFolderPath = newPath;

                using (var fileWriter = File.CreateText(_shapeRootFolderPathConfigFile))
                {
                    fileWriter.WriteLine(newPath);
                    fileWriter.Close();
                }

                MessageBox.Show(
                    string.Format(TextCollection.CustomeShapeSaveLocationChangedSuccessFormat, newPath),
                    TextCollection.CustomShapeSaveLocationChangedSuccessTitle, MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        private void DehighlightSelected()
        {
            if (_selectedThumbnail == null) return;
            
            _selectedThumbnail.DeHighlight();
            _selectedThumbnail = null;
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
            if (_selectedThumbnail != null)
            {
                if (_selectedThumbnail.State == LabeledThumbnail.Status.Editing)
                {
                    _selectedThumbnail.FinishNameEdit();
                }
                else
                    if (_selectedThumbnail == clickedThumbnail)
                    {
                        _clickOnSelected = true;
                    }

                _selectedThumbnail.DeHighlight();
            }

            clickedThumbnail.Highlight();

            _selectedThumbnail = clickedThumbnail;
        }

        private void FocusSelected()
        {
            myShapeFlowLayout.ScrollControlIntoView(_selectedThumbnail);
            _selectedThumbnail.Highlight();
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
            if (!FileAndDirTask.CopyFolder(oldPath, newPath))
            {
                loadingDialog.Dispose();

                MessageBox.Show(TextCollection.CustomShapeMigrationError);

                return false;
            }

            // now we will try our best to delete the original folder, but this is not guaranteed
            // because some of the using files, such as some opening shapes, and the evil thumb.db
            if (!FileAndDirTask.DeleteFolder(oldPath))
            {
                MessageBox.Show(TextCollection.CustomShapeOriginalFolderDeletionError);
            }

            ShapeRootFolderPath = newPath;

            // modify shape gallery presentation's path and name, then open it
            Globals.ThisAddIn.ShapePresentation.Path = newPath;
            Globals.ThisAddIn.ShapePresentation.ShapeFolderPath = CurrentShapeFolderPath;

            // if there's some lost during shape gallery opening, we must force reload the pane
            // to reflect the latest change
            //if (!Globals.ThisAddIn.ShapePresentation.Open(withWindow: false, focus: false))
            //{
            //    PaneReload(true);
            //}

            Globals.ThisAddIn.ShapePresentation.Open(withWindow: false, focus: false);
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

            DehighlightSelected();
        }

        private void RenameThumbnail(string oldName, LabeledThumbnail labeledThumbnail)
        {
            if (oldName == labeledThumbnail.NameLable) return;

            var newPath = labeledThumbnail.ImagePath.Replace(@"\" + oldName, @"\" + labeledThumbnail.NameLable);

            File.Move(labeledThumbnail.ImagePath, newPath);
            labeledThumbnail.ImagePath = newPath;

            Globals.ThisAddIn.ShapePresentation.RenameShape(oldName, labeledThumbnail.NameLable);

            Globals.ThisAddIn.SyncShapeRename(oldName, labeledThumbnail.NameLable);
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
        private void CustomShapePaneClick(object sender, EventArgs e)
        {
            if (_selectedThumbnail != null &&
                _selectedThumbnail.State == LabeledThumbnail.Status.Editing)
            {
                _selectedThumbnail.FinishNameEdit();
            }
        }

        private void FlowlayoutClick(object sender, EventArgs e)
        {
            if (_selectedThumbnail != null)
            {
                if (_selectedThumbnail.State == LabeledThumbnail.Status.Editing)
                {
                    _selectedThumbnail.FinishNameEdit();
                }
                else
                {
                    DehighlightSelected();
                }
            }
        }

        private void FlowlayoutContextMenuStripItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            var item = e.ClickedItem;

            if (item.Name.Contains("settings"))
            {
                ContextMenuStripSettingsClicked();
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
                Globals.ThisAddIn.ShapePresentation.CopyShape(shapeName);
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
                labeledThumbnail != _selectedThumbnail) return;

            // if name changed, rename the shape in shape gallery and the file on disk
            RenameThumbnail(oldName, labeledThumbnail);

            // put the labeled thumbnail to correct position
            ReorderThumbnail(labeledThumbnail);

            // select the thumbnail and scroll into view
            FocusSelected();
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
                _selectedThumbnail.StartNameEdit();
            }

            ClickTimerReset();
        }
        # endregion

        # region Comparer
        public class AtomicNumberStringCompare : IComparer<string>
        {
            public int Compare(string thisString, string otherString)
            {
                // some characters + number
                var pattern = new Regex(@"([^\d]+)(\d+)");
                var thisStringMatch = pattern.Match(thisString);
                var otherStringMatch = pattern.Match(otherString);

                // specially compare the pattern, after run out of the pattern, compare
                // 2 strings normally
                while (thisStringMatch.Success &&
                       otherStringMatch.Success)
                {
                    var thisStringPart = thisStringMatch.Groups[1].Value;
                    var thisNumPart = int.Parse(thisStringMatch.Groups[2].Value);

                    var otherStringPart = otherStringMatch.Groups[1].Value;
                    var otherNumPart = int.Parse(otherStringMatch.Groups[2].Value);

                    // if string part is not the same, we can tell the diff
                    if (!string.Equals(thisStringPart, otherStringPart))
                    {
                        break;
                    }

                    // if string part is the same but number part is different, we can
                    // tell the diff
                    if (thisNumPart != otherNumPart)
                    {
                        return thisNumPart - otherNumPart;
                    }

                    // two parts are identical, find next match
                    thisStringMatch = thisStringMatch.NextMatch();
                    otherStringMatch = otherStringMatch.NextMatch();
                }

                // case sensitive comparing, invariant for cultures
                return string.Compare(thisString, otherString, false, CultureInfo.InvariantCulture);
            }
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
