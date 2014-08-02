using System.IO;

namespace DeployHelper.Uploader
{
    class UploaderFactory
    {
        public static AbstractUploader GetUploader(DeployConfig config)
        {
            AbstractUploader uploader;
            switch (config.InstallerType)
            {
                case "online":
                    uploader = new OnlineInstallerUploader();
                    break;
                case "offline":
                    uploader = new OfflineInstallerUploader();
                    break;
                default:
                    Util.DisplayWarning("Invalid installer type found.", new InvalidDataException());
                    uploader = GetDummyPackager();
                    break;
            }
            uploader.SetConfig(config);
            return uploader;
        }

        private static AbstractUploader GetDummyPackager()
        {
            return new OnlineInstallerUploader();
        }
    }
}
