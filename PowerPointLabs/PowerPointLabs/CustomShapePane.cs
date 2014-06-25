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
        private readonly string _tempFolderName;
        private readonly string _tempFullPath;
        
        private bool _searchBoxFocused;

        public CustomShapePane(string tempFolderName)
        {
            InitializeComponent();

            _tempFolderName = @"\PowerPointLabs Temp\" + tempFolderName + @"\";
            _tempFullPath = Path.GetTempPath() + _tempFolderName;
            
            _searchBoxFocused = false;
        }

        private void panel1_DoubleClick(object sender, EventArgs e)
        {
            var tempPic = _tempFullPath + "temp.png";
            
            panel1.BackgroundImage.Save(tempPic);

            var currentSlide = PowerPointPresentation.CurrentSlide;
            var slideWidth = PowerPointPresentation.SlideWidth;
            var slideHeight = PowerPointPresentation.SlideHeight;

            if (currentSlide != null)
            {
                currentSlide.InsertPicture(tempPic, MsoTriState.msoFalse, MsoTriState.msoTrue, slideWidth/2,
                                           slideHeight/2);
            }

            File.Delete(tempPic);
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

        private void panel1_Enter(object sender, EventArgs e)
        {
            panel1.BackColor = Color.FromKnownColor(KnownColor.Highlight);
        }

        private void panel1_Leave(object sender, EventArgs e)
        {
            panel1.BackColor = Color.FromKnownColor(KnownColor.Control);
        }
    }
}
