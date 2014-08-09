using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PowerPointLabs.DataSources;
using PowerPointLabs.Utils;

namespace PowerPointLabs.Views
{
    public partial class ShapesLabSetting : Form
    {
        # region Properties and Bindings
        public Option UserOption { get; private set; }

        public string DefaultSavingPath { get; set; }

        private readonly ShapesLabSettingsDataSource _settingsDataSource = new ShapesLabSettingsDataSource();
        # endregion

        # region Constructors
        public ShapesLabSetting(string defaultPath)
        {
            InitializeComponent();

            _settingsDataSource.DefaultSavingPath = defaultPath;

            DataBindings.Add("DefaultSavingPath", _settingsDataSource, "DefaultSavingPath");

            pathBox.ReadOnly = true;
            pathBox.BackColor = Color.FromKnownColor(KnownColor.Window);
            pathBox.DataBindings.Add("Text", _settingsDataSource, "DefaultSavingPath");
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
                                       SelectedPath = _settingsDataSource.DefaultSavingPath,
                                       Description = TextCollection.FolderDialogDescription
                                   };

            var selectEmptyFolder = false;

            // loop until user chooses an empty folder, or click "Cancel" button
            while (!selectEmptyFolder)
            {
                // this launcher will scroll the view to selected path
                var result = FolderDialogLauncher.ShowFolderBrowser(folderDialog);

                if (result == DialogResult.OK)
                {
                    var newPath = folderDialog.SelectedPath;

                    if (!FileDir.IsDirectoryEmpty(newPath))
                    {
                        MessageBox.Show(TextCollection.FolderNonEmptyErrorMsg);
                    }
                    else
                    {
                        selectEmptyFolder = true;
                        _settingsDataSource.DefaultSavingPath = newPath;
                    }
                }
                else
                {
                    // if user cancel the dialog, break the loop
                    break;
                }
            }
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
    }
}
