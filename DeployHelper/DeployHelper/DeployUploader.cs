using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinSCP;

namespace DeployHelper
{
    class DeployUploader
    {
        #region SFTP upload

        //TODO: need refactor
        private static readonly string DirLocalPathToUpload = TextCollection.Config.DirCurrent + @"\PowerPointLabs_upload";
        private static readonly string DirAppFilesInLocalPath = DirLocalPathToUpload + @"\Application Files";
        private static string ZipPathToUpload = DirLocalPathToUpload + @"\PowerPointLabs.zip";
        private static string VstoPathToUpload = DirLocalPathToUpload + @"\PowerPointLabs.vsto";
        private readonly string DirPptLabsZipPath;

        private const Boolean IsToRemoveAfterUpload = true;

        private string _dirNameNewestVer = TextCollection.Config.DirBuildName;

        private string _releaseType;
        private string _installerType;

        public DeployUploader(string releaseType , string installerType)
        {
            _releaseType = releaseType;
            _installerType = installerType;
            if (_installerType == "offline")
            {
                DirPptLabsZipPath = TextCollection.Config.DirCurrent + @"\PowerPointLabs.zip";
            }
            else
            {
                DirPptLabsZipPath = TextCollection.Config.DirCurrent + "\\PowerPointLabsInstaller.zip";
            }
        }

        public void SftpUpload()
        {
            try
            {
                var sessionOptions = SetupSessionOptions();
                using (var session = new Session())
                {
                    //create a folder to upload and put PowerPointLabs.zip, PowerPointLabs.vsto,
                    //and Application Files/PowerPointLabs_X_X_X_X (the newest ver) into the folder
                    var dirNewestVerFolder = DirAppFilesInLocalPath + "\\" + _dirNameNewestVer;

                    Util.CreateDirectory(DirLocalPathToUpload);
                    if (_installerType == null || _installerType == "online")
                    {
                        Util.CreateDirectory(DirAppFilesInLocalPath);
                        Util.CreateDirectory(dirNewestVerFolder);
                    }

                    Console.WriteLine("Connecting the server...");
                    session.Open(sessionOptions);
                    if (session.Opened)
                    {
                        Util.DisplayDone(TextCollection.Const.DoneSftpConnected);
                        Console.WriteLine(TextCollection.Const.InfoFileUploading);

                        var remotePath = SetupRemotePath();
                        var transferOptions = SetupTransferOptions();
                        ConstructRemoteFolder(session, remotePath, transferOptions);

                        // Copy files into DirLocalPathToUpload
                        if (_installerType == null || _installerType == "online")
                        {
                            Util.CopyFolder(TextCollection.Config.DirBuild, dirNewestVerFolder, TextCollection.Const.IsOverWritten);
                            File.Copy(DirPptLabsZipPath, ZipPathToUpload, TextCollection.Const.IsOverWritten);
                            File.Copy(TextCollection.Config.DirVsto, VstoPathToUpload, TextCollection.Const.IsOverWritten);
                        }
                        else if (_installerType == "offline")
                        {
                            ZipPathToUpload = DirLocalPathToUpload + @"\PowerPointLabsInstaller.zip";
                            File.Copy(TextCollection.Config.DirCurrent + "\\PowerPointLabsInstaller.zip", ZipPathToUpload, TextCollection.Const.IsOverWritten);
                            VstoPathToUpload = DirLocalPathToUpload + @"\PowerPointLabsInstaller.vsto";
                            File.Copy(TextCollection.Config.DirVsto, VstoPathToUpload, TextCollection.Const.IsOverWritten);
                        }
                        File.Copy(TextCollection.Config.DirCurrent + "\\Tutorial.pptx", DirLocalPathToUpload + "\\Tutorial.pptx", TextCollection.Const.IsOverWritten);

                        Console.WriteLine("Uploading...");
                        UploadLocalFile(session, remotePath, transferOptions);
                        Util.DisplayDone(TextCollection.Const.DoneUploaded);
                    }
                    else
                    {
                        Util.DisplayWarning(TextCollection.Const.ErrorNetworkFailed, new InvalidOperationException());
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error during SFTP uploading: {0}", e);
                Util.DisplayWarning(TextCollection.Const.ErrorNetworkFailed, e);
            }
        }

        private static void UploadLocalFile(Session session, string remotePath, TransferOptions transferOptions)
        {
            var transferResult = session.PutFiles(
                DirLocalPathToUpload + @"\*",
                remotePath,
                !IsToRemoveAfterUpload,
                transferOptions);

            transferResult.Check();
        }

        private static void ConstructRemoteFolder(Session session, string remotePath, TransferOptions transferOptions)
        {
            // Construct folder with permissions first
            try
            {
                session.PutFiles(
                    DirLocalPathToUpload + @"\*",
                    remotePath,
                    !IsToRemoveAfterUpload,
                    transferOptions);
            }
            catch (InvalidOperationException)
            {
                if (session.Opened)
                {
                    Util.IgnoreException();
                }
                else
                {
                    throw;
                }
            }
        }

        private static TransferOptions SetupTransferOptions()
        {
            var transferOptions = new TransferOptions { TransferMode = TransferMode.Binary };
            var permissions = new FilePermissions { Octal = "644" };
            transferOptions.FilePermissions = permissions;
            return transferOptions;
        }

        private string SetupRemotePath()
        {
            string versionToUpload;
            if (_releaseType != null)
            {
                versionToUpload = _releaseType;
            }
            else
            {
                Console.Write(TextCollection.Const.InfoChooseVersion);
                versionToUpload = Console.ReadLine();
            }
            string remotePath = null;
            while (remotePath == null)
            {
                switch (versionToUpload)
                {
                    case TextCollection.Const.VarDev:
                        remotePath = TextCollection.Config.ConfigDevPath;
                        break;
                    case TextCollection.Const.VarRelease:
                        remotePath = TextCollection.Config.ConfigReleasePath;
                        break;
                    default:
                        Console.WriteLine("Incorrect release type!");
                        break;
                }
            }
            return remotePath;
        }

        //TODO: 1. hide server and username info, how? let user type once?
        private static SessionOptions SetupSessionOptions()
        {
            Console.Write(TextCollection.Const.InfoEnterPassword);
            var password = Console.ReadLine();
            while (password == null || password.Trim() == "")
            {
                Console.Write(TextCollection.Const.InfoEnterPassword);
                password = Console.ReadLine();
            }

            var sessionOptions = new SessionOptions
            {
                Protocol = Protocol.Sftp,
                HostName = TextCollection.Config.ConfigSftpAddress,
                UserName = TextCollection.Config.ConfigSftpUser,
                Password = password,
                PortNumber = 22, //TODO: make it configurable
                GiveUpSecurityAndAcceptAnySshHostKey = true
            };
            return sessionOptions;
        }

        #endregion
    }
}
