using System;
using System.IO;

namespace DeployHelper
{
    class InstallerPackager
    {
        #region Produce Zip

        //TODO: need refactor
        private string _dirCurrent = TextCollection.Config.DirCurrent;
        private string _dirVsto = TextCollection.Config.DirVsto;
        private string _dirNameNewestVer = TextCollection.Config.DirBuildName;
        private string _dirBuild = TextCollection.Config.DirBuild;
        private string _dirPptLabsZipPath;
        private string _installerType;

        public InstallerPackager(string installerType)
        {
            _installerType = installerType;
            _dirPptLabsZipPath = _dirCurrent + @"\PowerPointLabs.zip";
        }

        public void ProducePackage()
        {
            Console.WriteLine("Zipping...");
            switch (_installerType)
            {
                case "online":
                case null:
                {
                    var dirPptLabsZipFolder = _dirCurrent + @"\PowerPointLabs";
                    //rename setup.exe to bin.exe, and copy it into the data folder
                    var setupExeDirectory = _dirCurrent + @"\setup.exe";
                    var destSetupExeDirectory = _dirCurrent + @"\data\bin.exe";
                    File.Copy(setupExeDirectory, destSetupExeDirectory, TextCollection.Const.IsOverWritten);

                    //copy folder data, ReadMe.txt, and setup.bat
                    //into the powerPointLabsZipFolder; zip them together to produce PowerPointLabs.zip
                    //source:
                    var dirDataFolder = _dirCurrent + @"\data\";
                    var dirReadMe = _dirCurrent + @"\ReadMe.txt";
                    var dirSetupBat = _dirCurrent + @"\setup.bat";
                    //dest:
                    var dirZipDataFolder = dirPptLabsZipFolder + @"\data\";
                    var dirZipReadMe = dirPptLabsZipFolder + @"\ReadMe.txt";
                    var dirZipSetupBat = dirPptLabsZipFolder + @"\setup.bat";

                    //if dir not set up yet, then need to create them first
                    Util.CreateDirectory(dirPptLabsZipFolder);
                    Util.CreateDirectory(dirZipDataFolder);

                    try
                    {
                        Util.CopyFolder(dirDataFolder, dirZipDataFolder, TextCollection.Const.IsOverWritten);
                        File.Copy(dirReadMe, dirZipReadMe, TextCollection.Const.IsOverWritten);
                        File.Copy(dirSetupBat, dirZipSetupBat, TextCollection.Const.IsOverWritten);
                    }
                    catch (Exception e)
                    {
                        Util.DisplayWarning(TextCollection.Const.ErrorZipFilesMissing, e);
                    }

                    //remove the old zip file, if any
                    if (File.Exists(_dirPptLabsZipPath))
                    {
                        File.Delete(_dirPptLabsZipPath);
                    }
                    System.IO.Compression.ZipFile.CreateFromDirectory(dirPptLabsZipFolder, _dirPptLabsZipPath);
                }
                    break;
                case "offline":
                {
                    var dirPptLabsZipFolder = _dirCurrent + @"\PowerPointLabsInstaller";
                    //if dir not set up yet, then need to create them first
                    var dataFolder = dirPptLabsZipFolder + "\\data";
                    Util.CreateDirectory(dirPptLabsZipFolder);
                    Util.CreateDirectory(dataFolder);

                    try
                    {
                        File.Copy(_dirCurrent + "\\PowerPointLabsInstallerUi.exe",
                            dirPptLabsZipFolder + "\\setup.exe", TextCollection.Const.IsOverWritten);

                        File.Copy(_dirVsto, dataFolder + "\\PowerPointLabs.vsto", TextCollection.Const.IsOverWritten);
                        File.Copy(_dirCurrent + "\\setup.exe", dataFolder + "\\setup.exe", TextCollection.Const.IsOverWritten);

                        var dirAppFilesInLocalPath = dataFolder + "\\Application Files";
                        var dirNewestVerFolder = dirAppFilesInLocalPath + "\\" + _dirNameNewestVer;

                        Util.CreateDirectory(dirAppFilesInLocalPath);
                        Util.CreateDirectory(dirNewestVerFolder);
                        Util.CopyFolder(_dirBuild, dirNewestVerFolder, TextCollection.Const.IsOverWritten);

                        var dirPptLabsDataZipPath = dirPptLabsZipFolder + "\\data.zip";
                        var dirPptLabsInstallerZipPath = _dirCurrent + "\\PowerPointLabsInstaller.zip";
                        _dirPptLabsZipPath = dirPptLabsInstallerZipPath;
                        //remove the old zip file, if any
                        if (File.Exists(dirPptLabsDataZipPath))
                        {
                            File.Delete(dirPptLabsDataZipPath);
                        }
                        if (File.Exists(dirPptLabsInstallerZipPath))
                        {
                            File.Delete(dirPptLabsInstallerZipPath);
                        }
                        System.IO.Compression.ZipFile.CreateFromDirectory(dataFolder, dirPptLabsDataZipPath);
                        Directory.Delete(dataFolder, true);
                        System.IO.Compression.ZipFile.CreateFromDirectory(dirPptLabsZipFolder, dirPptLabsInstallerZipPath);
                    }
                    catch (Exception e)
                    {
                        Util.DisplayWarning(TextCollection.Const.ErrorZipFilesMissing, e);
                    }
                }
                    break;
                default:
                    Util.DisplayWarning("Invalid installer type found.", new InvalidDataException());
                    break;
            }
            Util.DisplayDone(TextCollection.Const.DoneZipped);
        }

        #endregion
    }
}
