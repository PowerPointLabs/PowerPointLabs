using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PowerPointLabs.Views
{
    public partial class ShapesLabSetting : Form
    {
        # region Properties
        public string DefaultShapeSavingPath { get; set; }
        
        public Option UserOption { get; private set; }
        # endregion

        # region Constructors
        public ShapesLabSetting(string defaultPath)
        {
            InitializeComponent();

            DefaultShapeSavingPath = defaultPath;

            pathBox.ReadOnly = true;
            pathBox.BackColor = Color.FromKnownColor(KnownColor.Window);
            pathBox.Text = DefaultShapeSavingPath;
        }
        # endregion

        # region API
        public enum Option
        {
            Ok,
            Cancel
        }
        # endregion

        # region Event Handlers
        private void BrowseButtonClick(object sender, EventArgs e)
        {
            var folderDialog = new FolderBrowserDialog
                                   {
                                       ShowNewFolderButton = true,
                                       SelectedPath = DefaultShapeSavingPath,
                                       Description = "Select the directory that you want to use as the default."
                                   };

            var result = FolderDialogLauncher.ShowFolderBrowser(folderDialog);
        }

        private void CancelButtonClick(object sender, EventArgs e)
        {
            UserOption = Option.Cancel;

            Dispose();
        }
        
        private void OkButtonClick(object sender, EventArgs e)
        {
            UserOption = Option.Ok;

            Dispose();
        }
        # endregion

        # region Helper Functions
        # endregion
    }
}
