using System.Windows.Forms;
using PowerPointLabs.Views;

namespace PowerPointLabs.Utils.Windows
{
    class FolderBrowserDialogUtil
    {
        public static string SelectFolder(string description, string selectedPath, bool showNewFolderButton = true)
        {
            return SelectFolderWinform(description, selectedPath, showNewFolderButton);
        }

        private static string SelectFolderWinform(string description, string selectedPath, bool showNewFolderButton)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog()
            {
                ShowNewFolderButton = showNewFolderButton,
                SelectedPath = selectedPath,
                Description = description
            };
            return (FolderDialogLauncher.ShowFolderBrowser(dialog) ==
                    System.Windows.Forms.DialogResult.OK)?
                dialog.SelectedPath : null;
        }
    }
}
