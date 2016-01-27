using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;

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
            const int waitTime = 1000;
            var tempFolder = _tempTestPath;
            var retryCount = 10;

            while (Directory.Exists(tempFolder) && retryCount > 0)
            {
                var tempFolderInfo = new DirectoryInfo(tempFolder);

                try
                {
                    DeepDeleteFolder(tempFolderInfo);
                }
                catch (Exception)
                {
                    retryCount--;
                    Thread.Sleep(waitTime);
                }
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
