using System;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Interop;
using System.Windows.Media.Imaging;

using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
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
            FolderBrowserDialog folderDialog = new FolderBrowserDialog
            {
                ShowNewFolderButton = true,
                SelectedPath = savePathInput.Text,
                Description = ShapesLabText.FolderDialogDescription
            };

            // loop until user chooses an empty folder, or click "Cancel" button
            while (true)
            {
                // this launcher will scroll the view to selected path
                System.Windows.Forms.DialogResult folderDialogResult = FolderDialogLauncher.ShowFolderBrowser(folderDialog);

                if (folderDialogResult == Forms.DialogResult.OK)
                {
                    string newPath = folderDialog.SelectedPath;

                    if (!FileDir.IsDirectoryEmpty(newPath))
                    {
                        System.Windows.MessageBox.Show(ShapesLabText.ErrorFolderNonEmpty);
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
