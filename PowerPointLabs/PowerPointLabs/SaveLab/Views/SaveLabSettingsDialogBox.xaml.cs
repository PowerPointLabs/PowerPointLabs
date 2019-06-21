using System.Windows;
using System.Windows.Forms;

using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;

using Forms = System.Windows.Forms;

namespace PowerPointLabs.SaveLab.Views
{
    /// <summary>
    /// Interaction logic for SaveLabSettingsDialogBox.xaml
    /// </summary>
    public partial class SaveLabSettingsDialogBox
    {
        public delegate void DialogConfirmedDelegate(string savePath);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

        public SaveLabSettingsDialogBox()
        {
            InitializeComponent();
        }
        
        public SaveLabSettingsDialogBox(string savePath)
            : this()
        {
            savePathBrowserIconImage.Source = GraphicsUtil.CreateBitmapSource(Properties.Resources.Load_icon);
            
            savePathInput.IsReadOnly = true;
            savePathInput.Text = savePath;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogConfirmedHandler(savePathInput.Text);
            Close();
        }

        private void SavePathBrowserButton_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog folderDialog = new FolderBrowserDialog
            {
                ShowNewFolderButton = true,
                SelectedPath = savePathInput.Text,
                Description = SaveLabText.FolderDialogDescription
            };

            // loop until user chooses a folder, or click "Cancel" button
            while (true)
            {
                // this launcher will scroll the view to selected path
                DialogResult folderDialogResult = FolderDialogLauncher.ShowFolderBrowser(folderDialog);

                if (folderDialogResult == Forms.DialogResult.OK)
                {
                    string newPath = folderDialog.SelectedPath;
                    savePathInput.Text = newPath;
                    break;
                }
                else
                {
                    // if user cancels the dialog, break the loop
                    break;
                }
            }
        }
    }
}
