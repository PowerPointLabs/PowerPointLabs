using System;
using System.IO;
using System.IO.Compression;

namespace DeployHelper.Packager
{
    class OfflineInstallerPackager : AbstractPackager
    {
        public readonly string DataFolderDirectory = DeployConfig.DirCurrent + @"\offline\data";

        public override void Produce()
        {
            //produce data.zip which offline installer's program needs
            try
            {
                //put required data into /offline folder
                Util.CreateDirectory(DataFolderDirectory);

                //data folder will look like this:
                //  \offline\data\PowerPointLabs.vsto
                //  \offline\data\setup.exe
                //  \offline\data\Application Files\PowerPointLabs_A_B_C_D\*
                File.Copy(DeployConfig.DirVsto, 
                    DataFolderDirectory + @"\PowerPointLabs.vsto", 
                    TextCollection.Const.IsOverWritten);
                File.Copy(DeployConfig.DirCurrent + @"\setup.exe",
                    DataFolderDirectory + @"\setup.exe", 
                    TextCollection.Const.IsOverWritten);

                var applicationFilesDirectory = DataFolderDirectory + @"\Application Files";
                var buildNameDirectory = DataFolderDirectory + @"\Application Files\" + Config.DirBuildName;

                Util.CreateDirectory(applicationFilesDirectory);
                Util.CreateDirectory(buildNameDirectory);
                Util.CopyFolder(Config.DirBuild, buildNameDirectory, TextCollection.Const.IsOverWritten);

                //make data folder into a zip file, named data.zip
                //  \offline\data.zip
                var dataZipPath = DeployConfig.DirOfflineInstallerFolder + @"\data.zip";
                //remove the old zip file, if any
                if (File.Exists(dataZipPath))
                {
                    File.Delete(dataZipPath);
                }
                ZipFile.CreateFromDirectory(DataFolderDirectory, dataZipPath);
                //remove data folder
                Directory.Delete(DataFolderDirectory, TextCollection.Const.IsSubDirectoryToDelete);
            }
            catch (Exception e)
            {
                Util.DisplayWarning(TextCollection.Const.ErrorZipFilesMissing, e);
            }

            //pack all
            //  \offline\data.zip
            //  \offline\setup.exe (offline installer program)
            //into this
            //  \offline\PowerPointLabsInstaller.exe (package, to be sent to user)
            RunIExpress(DeployConfig.DirOfflineInstallerFolder);
            //Turn IExpress package exe to zip
            if (File.Exists(DeployConfig.DirOfflineInstallerFolder + @"\PowerPointLabsInstaller.exe"))
            {
                var installerZip =
                    ZipStorer.Create(DeployConfig.DirOfflineInstallerFolder + @"\PowerPointLabsInstaller.zip",
                        "PowerPointLabs Offline Installer");
                installerZip.AddFile(ZipStorer.Compression.Store,
                    DeployConfig.DirOfflineInstallerFolder + @"\PowerPointLabsInstaller.exe"
                    , "setup.exe"
                    , "");
                installerZip.Close();
            }
            else
            {
                Util.DisplayWarning("IExpress package fails to produce", new FileNotFoundException());
            }
        }
    }
}
