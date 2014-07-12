using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPointLabs.Models;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;

namespace PowerPointLabs
{
    public partial class CustomShapePane : UserControl
    {
        private const string DefaultShapeNameFormat = @"My Shape Untitled {0}";
        private const string DefaultShapeNameSearchRegex = @"My Shape Untitled (\d+)";
        
        private LabeledThumbnail _selectedThumbnail;

        private bool _firstTimeLoading = true;

        private readonly AtomicNumberStringCompare _stringComparer = new AtomicNumberStringCompare();

        # region Properties
        public string NextDefaultFullName
        {
            get { return ShapeFolderPath + @"\" +
                         NextDefaultNameWithoutExtension + ".wmf"; }
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
            get { return ShapeFolderPath + @"\" +
                         CurrentShapeNameWithoutExtension + ".wmf"; }
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

        public string ShapeRootFolderPath { get; private set; }

        public string ShapeFolderPath
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

            ShowNoShapeMessage();
            myShapeFlowLayout.AutoSize = true;
            myShapeFlowLayout.Click += FlowlayoutClick;
        }
        # endregion

        # region API
        public void AddCustomShape(string shapeName, string shapeFullName, bool immediateEditing)
        {
            DehighlightSelected();

            var labeledThumbnail = new LabeledThumbnail(shapeFullName, shapeName);

            labeledThumbnail.ContextMenuStrip = contextMenuStrip;
            labeledThumbnail.Click += LabeledThumbnailClick;
            labeledThumbnail.DoubleClick += LabeledThumbnailDoubleClick;
            labeledThumbnail.NameChangedNotify += NameChangedNotifyHandler;

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
        }

        public void PaneReload()
        {
            if (!_firstTimeLoading)
            {
                return;
            }

            PrepareShapes();
            _firstTimeLoading = false;
        }
        # endregion

        # region Helper Functions
        private readonly Label _noShapeLabel = new Label
                                                   {
                                                       AutoSize = true,
                                                       Font =
                                                           new Font("Arial", 15.75F, FontStyle.Bold, GraphicsUnit.Point,
                                                                    0),
                                                       ForeColor = SystemColors.ButtonShadow,
                                                       Location = new Point(81, 11),
                                                       Name = "noShapeLabel",
                                                       Size = new Size(212, 24),
                                                       Text = TextCollection.CustomShapeNoShapeText
                                                   };

        private readonly Panel _noShapePanel = new Panel
                                                   {
                                                       Name = "noShapePanel",
                                                       Size = new Size(362, 50),
                                                       Margin = new Padding(0, 0, 0, 0)
                                                   };
        
        private void ContextMenuStripRemoveClicked()
        {
            if (_selectedThumbnail == null)
            {
                MessageBox.Show(TextCollection.CustomShapeNoPanelSelectedError);
                return;
            }

            File.Delete(CurrentShapeFullName);
            myShapeFlowLayout.Controls.Remove(_selectedThumbnail);
            _selectedThumbnail = null;

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

        private void DehighlightSelected()
        {
            if (_selectedThumbnail == null) return;
            
            _selectedThumbnail.DeHighlight();
            _selectedThumbnail = null;
        }

        private int FindControlIndex(string name)
        {
            if (myShapeFlowLayout.Controls.Contains(_noShapePanel))
            {
                return -1;
            }

            var totalControl = myShapeFlowLayout.Controls.Count;

            for (var i = 0; i < totalControl; i ++)
            {
                var control = myShapeFlowLayout.Controls[i] as LabeledThumbnail;
                
                if (control != null &&
                    _stringComparer.Compare(control.NameLable, name) >= 0)
                {
                    return i;
                }
            }

            return totalControl;
        }

        private void FocusSelected()
        {
            myShapeFlowLayout.ScrollControlIntoView(_selectedThumbnail);
            _selectedThumbnail.Highlight();
        }

        private void PrepareFolder()
        {
            if (!Directory.Exists(ShapeFolderPath))
            {
                Directory.CreateDirectory(ShapeFolderPath);
            }
        }

        private void PrepareShapes()
        {
            PrepareFolder();

            var shapes = Directory.EnumerateFiles(ShapeFolderPath, "*.png").OrderBy(item => item, _stringComparer);

            foreach (var shape in shapes)
            {
                var shapeName = Path.GetFileNameWithoutExtension(shape);

                if (shapeName == null)
                {
                    MessageBox.Show(TextCollection.CustomShapeWmfFileNameInvalid);
                    continue;
                }

                AddCustomShape(shapeName, shape, false);
            }

            DehighlightSelected();
        }

        private void ReorderThumbnail(LabeledThumbnail labeledThumbnail)
        {
            var index = FindControlIndex(labeledThumbnail.NameLable);

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
                _noShapePanel.Controls.Add(_noShapeLabel);
            }

            myShapeFlowLayout.Controls.Add(_noShapePanel);
        }

        private Tuple<Single, Single> ToMiddleOnScreen(Single slideWidth, Single slideHeight,
                                                       Single clientWidth, Single clientHeight)
        {
            return new Tuple<Single, Single>((slideWidth - clientWidth) / 2, (slideHeight - clientHeight) / 2);
        }
        # endregion

        # region Event Handlers

        private void ContextMenuStripItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            var item = e.ClickedItem;

            if (item.Name.Contains("remove"))
            {
                ContextMenuStripRemoveClicked();
            } else
            if (item.Name.Contains("edit"))
            {
                ContextMenuStripEditClicked();
            }
        }

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
            if (_selectedThumbnail != null &&
                _selectedThumbnail.State == LabeledThumbnail.Status.Editing)
            {
                _selectedThumbnail.FinishNameEdit();
            }
        }

        private void LabeledThumbnailClick(object sender, EventArgs e)
        {
            if (sender == null || !(sender is LabeledThumbnail))
            {
                MessageBox.Show(TextCollection.CustomShapeNoPanelSelectedError);
                return;
            }

            var clickedThumbnail = sender as LabeledThumbnail;

            if (_selectedThumbnail != null)
            {
                if (_selectedThumbnail.State == LabeledThumbnail.Status.Editing)
                {
                    _selectedThumbnail.FinishNameEdit();
                }

                _selectedThumbnail.DeHighlight();
            }

            clickedThumbnail.Highlight();
            _selectedThumbnail = clickedThumbnail;
        }

        private void LabeledThumbnailDoubleClick(object sender, EventArgs e)
        {
            if (sender == null || !(sender is LabeledThumbnail))
            {
                MessageBox.Show(TextCollection.CustomShapeNoPanelSelectedError);
                return;
            }

            var clickedThumbnail = sender as LabeledThumbnail;

            var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
            var image = clickedThumbnail.ImageToThumbnail;

            var slideWidth = PowerPointCurrentPresentationInfo.SlideWidth;
            var slideHeight = PowerPointCurrentPresentationInfo.SlideHeight;
            var clientWidth = (Single)image.Size.Width;
            var clientHeight = (Single)image.Size.Height;

            var leftTopCorner = ToMiddleOnScreen(slideWidth, slideHeight, clientWidth, clientHeight);

            if (currentSlide != null)
            {
                currentSlide.InsertPicture(clickedThumbnail.ImagePath, MsoTriState.msoFalse, MsoTriState.msoTrue,
                                           leftTopCorner);
            }
        }

        private void NameChangedNotifyHandler(object sender, bool nameChanged)
        {
            var labeledThumbnail = sender as LabeledThumbnail;

            // by right, name change only happens when the labeled thumbnail is selected.
            // Therfore, if the notifier doesn't come from the selected object, something
            // goes wrong.
            if (labeledThumbnail == null ||
                labeledThumbnail != _selectedThumbnail) return;

            ReorderThumbnail(labeledThumbnail);

            FocusSelected();
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
                    thisStringMatch.NextMatch();
                    otherStringMatch.NextMatch();
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
