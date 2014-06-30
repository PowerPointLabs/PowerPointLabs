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
        private const string DefaultShapeFolderName = @"\PowerPointLabs Custom Shapes";
        
        private LabeledThumbnail _selectedThumbnail;
        private int _currentUntitledShapeCnt = 0;
        private bool _firstTimeLoading = true;

        public string CurrentShapeName
        {
            get { return ShapeFolderPath + @"\" +
                         string.Format(DefaultShapeNameFormat, _currentUntitledShapeCnt) + ".wmf"; }
        }

        public string ShapeFolderPath
        {
            get { return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + DefaultShapeFolderName; }
        }

        public CustomShapePane()
        {
            InitializeComponent();

            int vertScrollWidth = SystemInformation.VerticalScrollBarWidth;

            myShapeFlowLayout.Padding = new Padding(0, 0, vertScrollWidth / 2, 0);
        }

        public void AddCustomShape()
        {
            var labeledThumbnail = new LabeledThumbnail();

            labeledThumbnail
        }

        public void PaneReload()
        {
            if (!_firstTimeLoading)
            {
                return;
            }
            
            //PrepareShapes();
            _firstTimeLoading = false;
        }

        # region Helper Functions
        private const string WmfFileNameInvalid = @"Invalid shape name encountered";

        private Tuple<Single, Single> ToMiddleOnScreen(Single slideWidth, Single slideHeight,
                                                       Single clientWidth, Single clientHeight)
        {
            return new Tuple<Single, Single>((slideWidth - clientWidth) / 2, (slideHeight - clientHeight) / 2);
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

                _currentUntitledShapeCnt = int.Parse(shapeName);

                AddCustomShape();
            }
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
            } else
            if (item.Name.Contains("edit"))
            {
                if (_selectedThumbnail == null)
                {
                    MessageBox.Show(NoPanelSelectedError);
                    return;
                }

                _selectedThumbnail.EnableNameEdit();
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

        private void OnNameLabelChanged(object sender, EventArgs e)
        {
            
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
