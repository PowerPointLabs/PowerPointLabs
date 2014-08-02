using System;
using System.IO;
using WinSCP;

namespace DeployHelper.Uploader
{
    abstract class AbstractUploader
    {
        protected DeployConfig Config;

        public void SetConfig(DeployConfig config)
        {
            Config = config;
        }

        public abstract void Prepare();

        public void CleanUp()
        {
            Directory.Delete(DeployConfig.DirLocalPathToUpload, TextCollection.Const.IsSubDirectoryToDelete);
        }

        public void UploadLocalFile(Session session, string remotePath, TransferOptions transferOptions)
        {
            ConstructRemoteFolder(session, remotePath, transferOptions);
            var transferResult = session.PutFiles(
                DeployConfig.DirLocalPathToUpload + @"\*",
                remotePath,
                !TextCollection.Const.IsToRemoveAfterUpload,
                transferOptions);

            transferResult.Check();
        }

        public void ConstructRemoteFolder(Session session, string remotePath, TransferOptions transferOptions)
        {
            // Construct folder with permissions first
            try
            {
                session.PutFiles(
                    DeployConfig.DirLocalPathToUpload + @"\*",
                    remotePath,
                    !TextCollection.Const.IsToRemoveAfterUpload,
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

        public TransferOptions SetupTransferOptions()
        {
            var transferOptions = new TransferOptions { TransferMode = TransferMode.Binary };
            var permissions = new FilePermissions { Octal = "644" };
            transferOptions.FilePermissions = permissions;
            return transferOptions;
        }

        public string SetupRemotePath()
        {
            var versionToUpload = Config.ReleaseType;
            switch (versionToUpload)
            {
                case TextCollection.Const.VarDev:
                    return Config.ConfigDevPath;
                case TextCollection.Const.VarRelease:
                    return Config.ConfigReleasePath;
                default:
                    Util.DisplayWarning("Incorrect release type!", new Exception());
                    //dummy return, won't reach
                    return Config.ConfigDevPath;
            }
        }

        //TODO: 1. hide server and username info, how? let user type once?
        public SessionOptions GetSessionOptions()
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
                HostName = Config.ConfigSftpAddress,
                UserName = Config.ConfigSftpUser,
                Password = password,
                PortNumber = 22, //TODO: make it configurable
                GiveUpSecurityAndAcceptAnySshHostKey = true
            };
            return sessionOptions;
        }
    }
}
