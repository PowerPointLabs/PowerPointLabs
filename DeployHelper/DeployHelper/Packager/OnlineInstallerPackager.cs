using System.IO;
using System.IO.Compression;

namespace DeployHelper.Packager
{
    class OnlineInstallerPackager : AbstractPackager
    {
        public override void Produce()
        {
            //put related data into /online folder
            var setupExeDirectory = DeployConfig.DirCurrent + @"\setup.exe";
            var destSetupExeDirectory = DeployConfig.DirOnlineInstallerFolder + @"\setup.exe";
            File.Copy(setupExeDirectory, destSetupExeDirectory, TextCollection.Const.IsOverWritten);

            //pack all
            //  \online\setup.exe (vsto's setup file)
            //  \online\setup.bat (batch file)
            //  \online\registry.reg (it's used to put SOC's website into IE's Trusted Sites)
            //into this
            //  \online\PowerPointLabs.exe  (package, to be sent to user)
            RunIExpress(DeployConfig.DirOnlineInstallerFolder);
            //Turn IExpress package exe to zip
            if (File.Exists(DeployConfig.DirOnlineInstallerFolder + @"\PowerPointLabs.exe"))
            {
                var installerZip =
                    ZipStorer.Create(DeployConfig.DirOnlineInstallerFolder + @"\PowerPointLabs.zip",
                        "PowerPointLabs Online Installer");
                installerZip.AddFile(ZipStorer.Compression.Store,
                    DeployConfig.DirOnlineInstallerFolder + @"\PowerPointLabs.exe"
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
