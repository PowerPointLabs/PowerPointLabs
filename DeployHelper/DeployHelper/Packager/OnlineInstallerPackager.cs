using System.IO;

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
        }
    }
}
