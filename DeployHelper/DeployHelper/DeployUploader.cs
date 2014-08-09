using System;
using DeployHelper.Uploader;
using WinSCP;

namespace DeployHelper
{
    class DeployUploader
    {
        #region SFTP upload

        private readonly DeployConfig _config;

        public DeployUploader(DeployConfig config)
        {
            _config = config;
        }

        public void SftpUpload()
        {
            try
            {
                using (var session = new Session())
                {
                    Console.WriteLine("Connecting the server...");
                    var uploader = UploaderFactory.GetUploader(_config);
                    session.Open(uploader.GetSessionOptions());
                    if (session.Opened)
                    {
                        Util.DisplayDone(TextCollection.Const.DoneSftpConnected);
                        Console.WriteLine(TextCollection.Const.InfoFileUploading);

                        var remotePath = uploader.SetupRemotePath();
                        var transferOptions = uploader.SetupTransferOptions();

                        uploader.Prepare();

                        Console.WriteLine("Uploading...");
                        uploader.UploadLocalFile(session, remotePath, transferOptions);

                        Util.DisplayDone(TextCollection.Const.DoneUploaded);
                        uploader.CleanUp();
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

        #endregion
    }
}
