using System.Windows;

using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Windows;

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
            // loop until user chooses a folder, or click "Cancel" button
            while (true)
            {
                string selectedPath = FolderBrowserDialogUtil.SelectFolder(
                    SaveLabText.FolderDialogDescription,
                    savePathInput.Text);

                if (selectedPath != null)
                {
                    string newPath = selectedPath;
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
