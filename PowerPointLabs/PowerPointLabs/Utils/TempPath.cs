using System.IO;

namespace PowerPointLabs.Utils
{
    public class TempPath
    {
        private static string _tempTestPath = Path.Combine(Path.GetTempPath(), "PowerPointLabsTest\\");

        public static string GetTempTestFolder()
        {
            return _tempTestPath;
        }

        public static bool IsExistingTempTestFolder()
        {
            return Directory.Exists(_tempTestPath);
        }

        #region Folder Operations
  
        public static void CreateTempTestFolder()
        {
            if (!IsExistingTempTestFolder())
            {
                Directory.CreateDirectory(_tempTestPath);
            }
        }

        public static void DeleteTempTestFolder()
        {
            string tempFolder = _tempTestPath;

            DirectoryInfo tempFolderInfo = new DirectoryInfo(tempFolder);

            try
            {
                DeepDeleteFolder(tempFolderInfo);
            }
            catch
            {
                // sometimes files cannot be deleted because in use
            }
        }

        private static void DeepDeleteFolder(DirectoryInfo rootFolder)
        {
            rootFolder.Attributes = FileAttributes.Normal;

            foreach (DirectoryInfo subFolder in rootFolder.GetDirectories())
            {
                DeepDeleteFolder(subFolder);
            }

            foreach (FileInfo file in rootFolder.GetFiles())
            {
                file.IsReadOnly = false;
                try
                {
                    file.Delete();
                }
                catch
                {
                    // sometimes files cannot be deleted because in use
                }
            }

            try
            {
                rootFolder.Delete(true);
            }
            catch
            {
                // sometimes the folder cannot be deleted
            }
        }

        #endregion

    }
}
