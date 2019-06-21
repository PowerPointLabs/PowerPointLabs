using System.Windows;

using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Windows;
using PowerPointLabs.Views;

using Forms = System.Windows.Forms;

namespace PowerPointLabs.ShapesLab.Views
{
    /// <summary>
    /// Interaction logic for ShapesLabSettingsDialogBox.xaml
    /// </summary>
    public partial class ShapesLabSettingsDialogBox
    {
        public delegate void DialogConfirmedDelegate(string savePath);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

        public ShapesLabSettingsDialogBox()
        {
            InitializeComponent();
        }
        
        public ShapesLabSettingsDialogBox(string savePath) : this()
        {
            savePathBrowserIconImage.Source = CommonUtil.CreateBitmapSource(Properties.Resources.Load_icon);
            
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
            Forms.FolderBrowserDialog folderDialog = new Forms.FolderBrowserDialog
            {
                ShowNewFolderButton = true,
                SelectedPath = savePathInput.Text,
                Description = ShapesLabText.FolderDialogDescription
            };

            // loop until user chooses an empty folder, or click "Cancel" button
            while (true)
            {
                // this launcher will scroll the view to selected path
                DialogResult folderDialogResult = FolderDialogLauncher.ShowFolderBrowser(folderDialog);

                if (folderDialogResult == Utils.Windows.DialogResult.OK)
                {
                    string newPath = folderDialog.SelectedPath;

                    if (!FileDir.IsDirectoryEmpty(newPath))
                    {
                        MessageBoxUtil.Show(ShapesLabText.ErrorFolderNonEmpty);
                    }
                    else
                    {
                        savePathInput.Text = newPath;
                        break;
                    }
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
