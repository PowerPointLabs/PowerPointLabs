using System.Windows;

using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Windows;

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
            savePathBrowserIconImage.Source = GraphicsUtil.BitmapToImageSource(Properties.Resources.Load_icon);
            
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
            // loop until user chooses an empty folder, or click "Cancel" button
            while (true)
            {
                // this launcher will scroll the view to selected path
                string selectedPath = FolderBrowserDialogUtil.SelectFolder(
                    ShapesLabText.FolderDialogDescription,
                    savePathInput.Text);

                if (selectedPath != null)
                {
                    string newPath = selectedPath;

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
