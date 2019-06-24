using System.Collections.Generic;
using System.Windows.Forms;

namespace PowerPointLabs.Utils.Windows
{
    class OpenFileDialogUtil
    {
        public static string Open(string title = "Open", string filter = "*", string defaultExt = "*")
        {
            return OpenWinform(title, filter, false, defaultExt)[0];
        }

        public static List<string> MultiOpen(string title = "Open", string filter = "*", string defaultExt = "*")
        {
            return OpenWinform(title, filter, true, defaultExt);
        }

        public static List<string> OpenWinform(string title, string filter, bool multiselect = false, string defaultExt = "*")
        {
            OpenFileDialog dialog = new OpenFileDialog()
            {
                Title = title,
                Filter = filter,
                Multiselect = multiselect,
                DefaultExt = defaultExt
            };
            return (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                ? new List<string>(dialog.FileNames) : null;
        }

    }
}
