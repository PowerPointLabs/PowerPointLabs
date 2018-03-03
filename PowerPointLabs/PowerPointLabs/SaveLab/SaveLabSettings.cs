using System.Drawing;
using System.IO;

using PowerPointLabs.SaveLab.Views;

namespace PowerPointLabs.SaveLab
{
    internal static class SaveLabSettings
    {
        public static string SaveFolderPath;

        private const string DefaultSaveMasterFolderName = @"\PowerPoint Save Lab Local Storage";
        private static readonly string DefaultSaveMasterFolderPrefix = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments);
        private static string defaultSavePath = DefaultSaveMasterFolderPrefix + DefaultSaveMasterFolderName;
        private static string defaultSaveTextFile = Path.Combine(defaultSavePath, "SavePath.txt");

        public static void ShowSettingsDialog()
        {
            SaveLabSettingsDialogBox dialog = new SaveLabSettingsDialogBox(SaveFolderPath);
            dialog.DialogConfirmedHandler += OnSettingsDialogConfirmed;
            dialog.ShowDialog();
        }

        // Function to be used at the start to create save directory and set initial value of path
        public static void InitialiseLocalStorage()
        {
            if (!Directory.Exists(defaultSavePath))
            {
                // Create the folder
                Directory.CreateDirectory(defaultSavePath);

                // Set initial value for SaveFolderPath
                SaveFolderPath = DefaultSaveMasterFolderPrefix;

                // Create a file to write to
                using (StreamWriter sw = File.CreateText(defaultSaveTextFile))
                {
                    sw.WriteLine(DefaultSaveMasterFolderPrefix.Trim());
                    sw.Close();
                }
            }
            else
            {
                // Read the SaveFolderPath from the local storage
                SaveFolderPath = ReadStoredPathStringFromLocalStorage();
            }
        }

        // Function updates the stpred path in the local storage with a new path
        private static void UpdateLocalStorage(string newPath)
        {
            // Overrides text with new path
            using (StreamWriter sw = new StreamWriter(defaultSaveTextFile, false))
            {
                sw.WriteLine(newPath.Trim());
                sw.Close();
            }
        }

        // Function updates the path string with the new path string
        private static void UpdatePathString(string newPath)
        {
            SaveFolderPath = newPath;
        }

        // Function reads the stored path from the text file in the local storage
        private static string ReadStoredPathStringFromLocalStorage()
        {
            // Read stored path string from existing text file
            using (StreamReader sr = new StreamReader(defaultSaveTextFile))
            {
                string storedPath = sr.ReadToEnd().Trim();
                sr.Close();
                return storedPath;
            }
        }

        private static void OnSettingsDialogConfirmed(string pathName)
        {
            UpdateLocalStorage(pathName);
            UpdatePathString(pathName);
            //SaveFolderPath = pathName;
        }
    }
}
