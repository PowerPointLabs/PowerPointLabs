using System;
using System.IO;
using PowerPointLabs.Properties;
using PowerPointLabs.Views;

namespace PowerPointLabs.ImageSearch.Util
{
    class TempPath
    {
        // resources & const
        public static string AggregatedTempFolder = Path.GetTempPath() + "pptlabs_imagesLab" + @"\";

        public static string AggregatedBackupTempFolder =
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\"
            + "pptlabs_imagesLab" + @"\";

        public static string TempFolder = AggregatedTempFolder + "pptlabs_imagesLab"
                                                   + DateTime.Now.GetHashCode() + @"\";
        public static readonly string BackupTempFolder = AggregatedBackupTempFolder + "pptlabs_imagesLab" 
            + DateTime.Now.GetHashCode() + @"\";

        public static readonly string LoadingImgPath = TempFolder + "loading_" + DateTime.Now.GetHashCode();
        public static readonly string LoadMoreImgPath = TempFolder + "loadMore_" + DateTime.Now.GetHashCode();

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
                    InitResources();
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
                    InitResources();
                }
                catch (Exception e)
                {
                    ErrorDialogWrapper.ShowDialog("Error", e.Message, e);
                    return false;
                }
            }
            return true;
        }

        private static void InitResources()
        {
            try
            {
                Resources.Loading.Save(LoadingImgPath);
                Resources.LoadMore.Save(LoadMoreImgPath);
            }
            catch
            {
                // may fail to save it, which is fine
            }
        }

        public static string GetPath(string name)
        {
            var fullsizeImageFile = TempFolder + name + "_"
                                    + Guid.NewGuid().ToString().Substring(0, 6)
                                    + DateTime.Now.GetHashCode();
            return fullsizeImageFile;
        }

        public static void Empty(DirectoryInfo directory)
        {
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
