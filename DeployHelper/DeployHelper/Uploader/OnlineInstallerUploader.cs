using System.IO;

namespace DeployHelper.Uploader
{
    class OnlineInstallerUploader : AbstractUploader
    {
        public override void Prepare()
        {
            //create a folder to upload and put PowerPointLabs.exe, PowerPointLabs.vsto,
            //and Application Files/PowerPointLabs_X_X_X_X (the newest ver) into the upload folder
            Util.CreateDirectory(DeployConfig.DirLocalPathToUpload);
            Util.CreateDirectory(DeployConfig.DirAppFilesToUpload);
            var buildNameDirectory = DeployConfig.DirAppFilesToUpload + @"\" + Config.DirBuildName;
            Util.CreateDirectory(buildNameDirectory);

            //build folder
            Util.CopyFolder(Config.DirBuild, buildNameDirectory, TextCollection.Const.IsOverWritten);
            //installer exe zip
            File.Copy(DeployConfig.DirOnlineInstallerFolder + @"\PowerPointLabs.zip",
                DeployConfig.DirLocalPathToUpload + @"\PowerPointLabs.zip", 
                TextCollection.Const.IsOverWritten);
            //vsto file
            File.Copy(DeployConfig.DirVsto,
                DeployConfig.DirLocalPathToUpload + @"\PowerPointLabs.vsto", 
                TextCollection.Const.IsOverWritten);
            //tutorial file
            File.Copy(DeployConfig.DirOnlineInstallerFolder + @"\Tutorial.pptx", 
                DeployConfig.DirLocalPathToUpload + @"\Tutorial.pptx", 
                TextCollection.Const.IsOverWritten);
        }
    }
}
