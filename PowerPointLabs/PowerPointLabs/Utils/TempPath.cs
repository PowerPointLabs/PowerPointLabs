using System;
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

        public static bool IsExistingTempFolder()
        {
            return Directory.Exists(_tempTestPath);
        }

        #region Folder Operations
  
        public static void CreateTempTestFolder()
        {
            Directory.CreateDirectory(_tempTestPath);
        }

        public static void DeleteTempTestFolder()
        {
            var tempFolder = _tempTestPath;

            while (Directory.Exists(tempFolder))
            {
                var tempFolderInfo = new DirectoryInfo(tempFolder);

                try
                {
                    DeepDeleteFolder(tempFolderInfo);
                }
                catch (IOException)
                { }
            }
        }

        private static void DeepDeleteFolder(DirectoryInfo rootFolder)
        {
            rootFolder.Attributes = FileAttributes.Normal;

            foreach (var subFolder in rootFolder.GetDirectories())
            {
                DeepDeleteFolder(subFolder);
            }

            foreach (var file in rootFolder.GetFiles())
            {
                file.IsReadOnly = false;
            }

            rootFolder.Delete(true);
        }

        #endregion

    }
}
