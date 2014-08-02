using System.IO;

namespace DeployHelper.Uploader
{
    class OfflineInstallerUploader : AbstractUploader
    {
        public override void Prepare()
        {
            Util.CreateDirectory(DeployConfig.DirLocalPathToUpload);
            //offline installer exe
            File.Copy(DeployConfig.DirOfflineInstallerFolder + @"\PowerPointLabsInstaller.exe",
                DeployConfig.DirLocalPathToUpload + @"\PowerPointLabsInstaller.exe", 
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
