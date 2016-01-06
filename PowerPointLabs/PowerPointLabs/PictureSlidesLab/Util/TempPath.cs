using System;
using System.IO;
using PowerPointLabs.Views;

namespace PowerPointLabs.PictureSlidesLab.Util
{
    class TempPath
    {
        // resources & const
        public static string AggregatedTempFolder = Path.GetTempPath() + "pptlabs_pictureSlidesLab" + @"\";

        public static string AggregatedBackupTempFolder =
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\"
            + "pptlabs_pictureSlidesLab" + @"\";

        public static string TempFolder = AggregatedTempFolder + "pptlabs_pictureSlidesLab"
                                                   + DateTime.Now.GetHashCode() + @"\";
        public static readonly string BackupTempFolder = AggregatedBackupTempFolder + "pptlabs_pictureSlidesLab"
            + DateTime.Now.GetHashCode() + @"\";

        /// <returns>is successful</returns>
        public static bool InitTempFolder()
        {
            return Init() || RetryInit();
        }

        private static bool Init()
        {
            Empty(new DirectoryInfo(AggregatedTempFolder));
            if (!Directory.Exists(TempFolder))
            {
                try
                {
                    Directory.CreateDirectory(TempFolder);
                }
                catch
                {
                    return false;
                }
            }
            return true;
        }

        private static bool RetryInit()
        {
            TempFolder = BackupTempFolder;
            Empty(new DirectoryInfo(AggregatedBackupTempFolder));
            if (!Directory.Exists(TempFolder))
            {
                try
                {
                    Directory.CreateDirectory(TempFolder);
                }
                catch (Exception e)
                {
                    ErrorDialogWrapper.ShowDialog("Error", e.Message, e);
                    return false;
                }
            }
            return true;
        }

        public static string GetPath(string name)
        {
            var fullsizeImageFile = TempFolder + name + "_"
                                    + Guid.NewGuid().ToString().Substring(0, 6)
                                    + DateTime.Now.GetHashCode();
            return fullsizeImageFile;
        }

        private static void Empty(DirectoryInfo directory)
        {
            if (!directory.Exists) return;

            try
            {
                foreach (FileInfo file in directory.GetFiles()) file.Delete();
                foreach (DirectoryInfo subDirectory in directory.GetDirectories()) subDirectory.Delete(true);
            }
            catch (Exception)
            {
                // ignore ex, if cannot delete trash
            }
        }
    }
}
