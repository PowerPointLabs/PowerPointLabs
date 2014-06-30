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
        private LabeldThumbnail _selectedPanel;
        private int _currentShapeCnt = 0;
        private bool _firstTimeLoading = true;

        public string CurrentShapeName
        {
            get { return ShapeFolderPath + @"\" + _currentShapeCnt + ".wmf"; }
        }

        public string ShapeFolderPath
        {
            get { return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\PowerPointLabs Custom Shapes"; }
        }

        public CustomShapePane()
        {
            InitializeComponent();

            int vertScrollWidth = SystemInformation.VerticalScrollBarWidth;

            myShapeFlowLayout.Padding = new Padding(0, 0, vertScrollWidth / 2, 0);
        }

        public void AddCustomShape()
        {
            var shapeImage = new Bitmap(CurrentShapeName);
            
            var newShapeCell = new Panel();

            newShapeCell.Size = new Size(50, 50);
            newShapeCell.Name = CurrentShapeName;
            //newShapeCell.BackgroundImage = CreateThumbnailImage(shapeImage, 50, 50);
            newShapeCell.ContextMenuStrip = contextMenuStrip;
            newShapeCell.DoubleClick += PanelDoubleClick;
            newShapeCell.Click += PanelClick;

            myShapeFlowLayout.Controls.Add(newShapeCell);

            if ((myShapeFlowLayout.Controls.Count + 1) * 56 < motherTableLayoutPanel.Size.Width)
            {
                myShapeFlowLayout.AutoSize = false;
            }
            else
            {
                myShapeFlowLayout.AutoSize = true;
            }

            _currentShapeCnt++;
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

                _currentShapeCnt = int.Parse(shapeName);

                AddCustomShape();
            }
        }
        # endregion

        # region Event Handlers
        private const string NoPanelSelectedError = @"No shape selected";

        private void PanelDoubleClick(object sender, EventArgs e)
        {
            if (sender == null || !(sender is Panel))
            {
                MessageBox.Show(NoPanelSelectedError);
                return;
            }

            var childPanel = sender as Panel;

            var currentSlide = PowerPointPresentation.CurrentSlide;
            var image = new Bitmap(childPanel.Name);
            
            var slideWidth = PowerPointPresentation.SlideWidth;
            var slideHeight = PowerPointPresentation.SlideHeight;
            var clientWidth = (Single)image.Size.Width;
            var clientHeight = (Single)image.Size.Height;

            var leftTopCorner = ToMiddleOnScreen(slideWidth, slideHeight, clientWidth, clientHeight);

            if (currentSlide != null)
            {
                currentSlide.InsertPicture(childPanel.Name, MsoTriState.msoFalse, MsoTriState.msoTrue, leftTopCorner);
            }
        }

        private void PanelClick(object sender, EventArgs e)
        {
            if (sender == null || !(sender is Panel))
            {
                MessageBox.Show(NoPanelSelectedError);
                return;
            }

            var childPanel = sender as Panel;

            // de-highlight the old shape and set current shape as highighted
            if (_selectedPanel != null)
            {
                _selectedPanel.BackColor = Color.Transparent;
            }

            childPanel.BackColor = Color.FromKnownColor(KnownColor.Highlight);
            _selectedPanel = childPanel;
        }

        private void ContextMenuStripItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            var item = e.ClickedItem;

            if (!item.Name.Contains("remove"))
            {
                return;
            }

            if (_selectedPanel == null)
            {
                MessageBox.Show(NoPanelSelectedError);
                return;
            }

            File.Delete(_selectedPanel.Name);
            myShapeFlowLayout.Controls.Remove(_selectedPanel);
            _selectedPanel = null;
        }
        # endregion

        private void LabeldThumbnailClick(object sender, EventArgs e)
        {
            if (sender == null || !(sender is LabeldThumbnail))
            {
                MessageBox.Show(NoPanelSelectedError);
                return;
            }

            var clickedThumbnail = sender as LabeldThumbnail;

            if (_selectedPanel != null)
            {
                _selectedPanel = clickedThumbnail;
            }
        }

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
