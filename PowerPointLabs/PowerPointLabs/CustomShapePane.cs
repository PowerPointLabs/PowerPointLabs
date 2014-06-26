using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PPExtraEventHelper;
using PowerPointLabs.Models;

namespace PowerPointLabs
{
    public partial class CustomShapePane : UserControl
    {
        private bool _searchBoxFocused;
        private Panel _selectedPanel;

        public CustomShapePane()
        {
            InitializeComponent();

            _searchBoxFocused = false;
        }

        public void AddCustomShape(string fileName)
        {
            var shapeImage = new Bitmap(fileName);
            
            var newShapeCell = new Panel();

            newShapeCell.Size = new Size(50, 50);
            newShapeCell.Name = fileName;
            newShapeCell.BackgroundImage = CreateThumbnailImage(shapeImage, 50, 50);
            newShapeCell.ContextMenuStrip = contextMenuStrip;
            newShapeCell.DoubleClick += PanelDoubleClick;
            newShapeCell.Click += PanelClick;

            myShapeFlowLayout.Controls.Add(newShapeCell);
        }

        private Tuple<Single, Single> ToMiddleOnScreen(Single slideWidth, Single slideHeight,
                                                       Single clientWidth, Single clientHeight)
        {
            return new Tuple<Single, Single>((slideWidth - clientWidth) / 2, (slideHeight - clientHeight) / 2);
        }

        private double CalculateScalingRatio(Size oldSize, Size newSize)
        {
            double scalingRatio;

            if (oldSize.Width >= oldSize.Height)
            {
                scalingRatio = (double)newSize.Width / oldSize.Width;
            }
            else
            {
                scalingRatio = (double)newSize.Height / oldSize.Height;
            }

            return scalingRatio;
        }

        private Bitmap CreateThumbnailImage(Image oriImage, int width, int height)
        {
            var scalingRatio = CalculateScalingRatio(oriImage.Size, new Size(width, height));

            // calculate width and height after scaling
            var scaledWidth = (int)Math.Round(oriImage.Size.Width * scalingRatio);
            var scaledHeight = (int)Math.Round(oriImage.Size.Height * scalingRatio);

            // calculate left top corner position of the image in the thumbnail
            var scaledLeft = (width - scaledWidth) / 2;
            var scaledTop = (height - scaledHeight) / 2;

            // define drawing area
            var drawingRect = new Rectangle(scaledLeft, scaledTop, scaledWidth, scaledHeight);
            var thumbnail = new Bitmap(width, height);

            // here we set the thumbnail as the highest quality
            using (var thumbnailGraphics = Graphics.FromImage(thumbnail))
            {
                thumbnailGraphics.CompositingQuality = CompositingQuality.HighQuality;
                thumbnailGraphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                thumbnailGraphics.SmoothingMode = SmoothingMode.HighQuality;
                thumbnailGraphics.DrawImage(oriImage, drawingRect);
            }

            return thumbnail;
        }

        # region Event Handlers
        private const string PanelNullClickSenderError = @"No shape selected";

        private void PanelDoubleClick(object sender, EventArgs e)
        {
            if (sender == null || !(sender is Panel))
            {
                MessageBox.Show(PanelNullClickSenderError);
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
                MessageBox.Show(PanelNullClickSenderError);
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

            myShapeFlowLayout.Controls.Remove(_selectedPanel);
            _selectedPanel = null;
        }
        # endregion

        # region search box appearance and behaviors
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
        # endregion
    }
}
