using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PPExtraEventHelper;
using PowerPointLabs.Models;

namespace PowerPointLabs
{
    public partial class CustomShapePane : UserControl
    {
        private const string DefaultShapeNameFormat = @"My Shape Untitled {0}";
        private const string DefaultShapeNameSearchRegex = @"My Shape Untitled (\d+)";
        private const string DefaultShapeFolderName = @"\PowerPointLabs Custom Shapes\My Shapes";
        
        private LabeledThumbnail _selectedThumbnail;

        private bool _firstTimeLoading = true;

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
                var min = myShapeFlowLayout.Controls.Count;
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

                return string.Format(DefaultShapeNameFormat, min + 1);
            }
        }

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

        public string ShapeFolderPath
        {
            get { return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + DefaultShapeFolderName; }
        }
        # endregion

        # region Constructors
        public CustomShapePane()
        {
            InitializeComponent();

            ShowNoShapeMessage();
            //myShapeFlowLayout.AutoSize = true;
            myShapeFlowLayout.Click += FlowlayoutClick;

            var vertScrollWidth = SystemInformation.VerticalScrollBarWidth;
            myShapeFlowLayout.Padding = new Padding(0, 0, vertScrollWidth / 2, 0);
        }
        # endregion

        # region API
        public void AddCustomShape(string shapeName, string shapeFullName, bool immediateEditing)
        {
            // remove no_shape banner if we have one
            if (myShapeFlowLayout.Controls.Contains(_noShapePanel))
            {
                myShapeFlowLayout.Controls.Remove(_noShapePanel);
            }

            // dehighlight the selected labeled thumbnail if we have one
            if (_selectedThumbnail != null)
            {
                _selectedThumbnail.DeHighlight();
            }

            var labeledThumbnail = new LabeledThumbnail(shapeFullName, shapeName);

            labeledThumbnail.ContextMenuStrip = contextMenuStrip;
            labeledThumbnail.Click += LabeledThumbnailClick;
            labeledThumbnail.DoubleClick += LabeledThumbnailDoubleClick;
            labeledThumbnail.NameChangedNotify += NameChangedNotifyHandler;

            myShapeFlowLayout.Controls.Add(labeledThumbnail);
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
        private const string WmfFileNameInvalid = @"Invalid shape name encountered";
        private const string NoShapeText = @"No shapes available";

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
                                                       Text = NoShapeText
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
                MessageBox.Show(NoPanelSelectedError);
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
                MessageBox.Show(NoPanelSelectedError);
                return;
            }

            _selectedThumbnail.StartNameEdit();
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

            var wmfFiles = Directory.EnumerateFiles(ShapeFolderPath, "*.wmf");

            foreach (var wmfFile in wmfFiles)
            {
                var shapeName = Path.GetFileNameWithoutExtension(wmfFile);

                if (shapeName == null)
                {
                    MessageBox.Show(WmfFileNameInvalid);
                    continue;
                }

                AddCustomShape(shapeName, wmfFile, false);
            }
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
        private const string NoPanelSelectedError = @"No shape selected";

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
                MessageBox.Show(NoPanelSelectedError);
                return;
            }

            var clickedThumbnail = sender as LabeledThumbnail;

            if (_selectedThumbnail != null)
            {
                if (_selectedThumbnail.State == LabeledThumbnail.Status.Editing)
                {
                    _selectedThumbnail.FinishNameEdit();
                }
                else
                {
                    _selectedThumbnail.DeHighlight();
                }
            }

            clickedThumbnail.Highlight();
            _selectedThumbnail = clickedThumbnail;
        }

        private void LabeledThumbnailDoubleClick(object sender, EventArgs e)
        {
            if (sender == null || !(sender is LabeledThumbnail))
            {
                MessageBox.Show(NoPanelSelectedError);
                return;
            }

            var clickedThumbnail = sender as LabeledThumbnail;

            var currentSlide = PowerPointPresentation.CurrentSlide;
            var image = clickedThumbnail.ImageToThumbnail;

            var slideWidth = PowerPointPresentation.SlideWidth;
            var slideHeight = PowerPointPresentation.SlideHeight;
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

            labeledThumbnail.DeHighlight();
            _selectedThumbnail = null;
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
