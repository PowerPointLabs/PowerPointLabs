using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PPExtraEventHelper;
using PowerPointLabs.Models;
using Stepi.UI;

namespace PowerPointLabs
{
    public partial class CustomShapePane : UserControl
    {
        private bool _searchBoxFocused;
        private List<string> _myShapes;
        private Panel _selectedPanel;

        public CustomShapePane()
        {
            InitializeComponent();

            _myShapes = new List<string>();
            
            _searchBoxFocused = false;
        }

        private bool ThumbNailCallBack()
        {
            return false;
        }

        public void AddCustomShape(string fileName)
        {
            _myShapes.Add(fileName);

            var shapeImage = new Bitmap(fileName);
            
            var newShapeCell = new Panel();

            newShapeCell.Size = new Size(50, 50);
            newShapeCell.Name = fileName;
            newShapeCell.BackgroundImage = shapeImage.GetThumbnailImage(50, 50, ThumbNailCallBack,
                                                                        IntPtr.Zero);
            newShapeCell.DoubleClick += PanelDoubleClick;
            newShapeCell.Click += PanelClick;

            myShapeFlowLayout.Controls.Add(newShapeCell);
        }

        private void PanelDoubleClick(object sender, EventArgs e)
        {
            var childPanel = sender as Panel;

            var currentSlide = PowerPointPresentation.CurrentSlide;
            var slideWidth = PowerPointPresentation.SlideWidth;
            var slideHeight = PowerPointPresentation.SlideHeight;

            if (currentSlide != null)
            {
                currentSlide.InsertPicture(childPanel.Name, MsoTriState.msoFalse, MsoTriState.msoTrue, slideWidth/2,
                                           slideHeight/2);
            }
        }

        private void PanelClick(object sender, EventArgs e)
        {
            var childPanel = sender as Panel;

            // de-highlight the old shape and set current shape as highighted
            if (_selectedPanel != null)
            {
                _selectedPanel.BackColor = Color.Transparent;
            }

            childPanel.BackColor = Color.FromKnownColor(KnownColor.Highlight);
            _selectedPanel = childPanel;
        }

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
