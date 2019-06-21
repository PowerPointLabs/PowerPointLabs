using System.Windows.Forms;

namespace PowerPointLabs.Utils.Windows
{
    class SaveFileDialogUtil
    {
        public static string? Save(string title = "Save As", string filter = "*", string defaultExt = "",  string filename="", string initialDirectory = "", bool overwriteprompt = true)
        {
            return SaveWinForm(defaultExt, filter, title, filename, initialDirectory, overwriteprompt);
        }

        public static string? SaveWinForm(string defaultExt, string filter,
            string title, string filename,
            string initialDirectory, bool overwriteprompt)
        {
            SaveFileDialog dialog = new SaveFileDialog()
            {
                DefaultExt = defaultExt,
                Filter = filter,
                Title = title,
                FileName = filename,
                InitialDirectory = initialDirectory,
                OverwritePrompt = overwriteprompt
            };
            return (DialogResult)(int)dialog.ShowDialog();
        }
    }
}
