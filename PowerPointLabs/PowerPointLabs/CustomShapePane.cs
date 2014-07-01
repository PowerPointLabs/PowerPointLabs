using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PPExtraEventHelper;
using PowerPointLabs.Models;

namespace PowerPointLabs
{
    public partial class CustomShapePane : UserControl
    {
        private const string DefaultShapeNameFormat = @"My Shape Untitled {0}";
        private const string DefaultShapeFolderName = @"\PowerPointLabs Custom Shapes\My Shapes";
        
        private LabeledThumbnail _selectedThumbnail;

        private int _currentUntitledShapeCnt = 0;
        private bool _firstTimeLoading = true;

        # region Properties
        public string CurrentShapeFullName
        {
            get { return ShapeFolderPath + @"\" +
                         CurrentShapeNameWithoutExtension + ".wmf"; }
        }

        public string CurrentShapeNameWithoutExtension
        {
            get { return string.Format(DefaultShapeNameFormat, _currentUntitledShapeCnt); }
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

            int vertScrollWidth = SystemInformation.VerticalScrollBarWidth;

            myShapeFlowLayout.Padding = new Padding(0, 0, vertScrollWidth / 2, 0);
        }
        # endregion

        # region API
        public void AddCustomShape()
        {
            if (myShapeFlowLayout.Controls.Contains(_noShapePanel))
            {
                myShapeFlowLayout.Controls.Remove(_noShapePanel);
            }

            var labeledThumbnail = new LabeledThumbnail(CurrentShapeFullName, CurrentShapeNameWithoutExtension);

            labeledThumbnail.Click += LabeledThumbnailClick;
            labeledThumbnail.DoubleClick += LabeledThumbnailDoubleClick;

            myShapeFlowLayout.Controls.Add(labeledThumbnail);
            myShapeFlowLayout.ScrollControlIntoView(labeledThumbnail);

            labeledThumbnail.StartNameEdit();
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
                                                       Size = new Size(362, 46)
                                                   };

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

                _currentUntitledShapeCnt = int.Parse(shapeName);

                AddCustomShape();
            }

            if (myShapeFlowLayout.Controls.Count == 0)
            {
                ShowNoShapeMessage();
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
                if (_selectedThumbnail == null)
                {
                    MessageBox.Show(NoPanelSelectedError);
                    return;
                }

                File.Delete(_selectedThumbnail.Name);
                myShapeFlowLayout.Controls.Remove(_selectedThumbnail);
                _selectedThumbnail = null;

                if (myShapeFlowLayout.Controls.Count == 0)
                {
                    ShowNoShapeMessage();
                }
            } else
            if (item.Name.Contains("edit"))
            {
                if (_selectedThumbnail == null)
                {
                    MessageBox.Show(NoPanelSelectedError);
                    return;
                }

                _selectedThumbnail.StartNameEdit();
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
                _selectedThumbnail.ToggleHighlight();
            }

            clickedThumbnail.ToggleHighlight();
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
