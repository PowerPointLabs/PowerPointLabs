using System.IO;

namespace DeployHelper.Uploader
{
    class OfflineInstallerUploader : AbstractUploader
    {
        public override void Prepare()
        {
            Util.CreateDirectory(DeployConfig.DirLocalPathToUpload);
            //offline installer exe zip
            File.Copy(DeployConfig.DirOfflineInstallerFolder + @"\PowerPointLabsInstaller.zip",
                DeployConfig.DirLocalPathToUpload + @"\PowerPointLabsInstaller.zip", 
                TextCollection.Const.IsOverWritten);
            //offline installer data zip
            File.Copy(DeployConfig.DirOfflineInstallerFolder + @"\data.zip",
                DeployConfig.DirLocalPathToUpload + @"\data.zip",
                TextCollection.Const.IsOverWritten);
            //vsto file
            File.Copy(DeployConfig.DirVsto, 
                DeployConfig.DirLocalPathToUpload + @"\PowerPointLabsInstaller.vsto", 
                TextCollection.Const.IsOverWritten);
            //tutorial file
            File.Copy(DeployConfig.DirOfflineInstallerFolder + @"\Tutorial.pptx",
                DeployConfig.DirLocalPathToUpload + @"\Tutorial.pptx", 
                TextCollection.Const.IsOverWritten);
        }
    }
}
