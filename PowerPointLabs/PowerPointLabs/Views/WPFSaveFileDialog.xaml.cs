using System.Windows.Controls;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Windows;

namespace PowerPointLabs.Views
{
    /// <summary>
    /// Interaction logic for WPFSaveDialogFile.xaml
    /// </summary>
    public partial class WPFSaveFileDialog
    {
        public string DefaultExt { get; set; }
        public string Filter { get; set; }
        public string Title { get; set; }
        public string FileName { get; set; }
        public string InitialDirectory { get; set; }
        public bool OverwritePrompt { get; set; }

        public WPFSaveFileDialog()
        {
            InitializeComponent();
        }

        public DialogResult ShowDialog()
        {
            return DialogResult.None;
        }
    }
}
