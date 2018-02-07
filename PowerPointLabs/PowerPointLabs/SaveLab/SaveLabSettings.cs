using System.Drawing;

using PowerPointLabs.SaveLab.Views;

namespace PowerPointLabs.SaveLab
{
    internal static class SaveLabSettings
    {
        public static string SaveFolderPath = "";

        public static void ShowSettingsDialog()
        {
            SaveLabSettingsDialogBox dialog = dialog = new SaveLabSettingsDialogBox(SaveFolderPath);
            dialog.DialogConfirmedHandler += OnSettingsDialogConfirmed;
            dialog.ShowDialog();
        }

        private static void OnSettingsDialogConfirmed(string pathName)
        {
            SaveFolderPath = pathName;
        }
    }
}
