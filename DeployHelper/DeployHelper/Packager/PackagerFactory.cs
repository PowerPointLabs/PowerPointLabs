using System.IO;

namespace DeployHelper.Packager
{
    class PackagerFactory
    {
        public static AbstractPackager GetPackager(DeployConfig config)
        {
            AbstractPackager packager;
            switch (config.InstallerType)
            {
                case "online":
                    packager = new OnlineInstallerPackager();
                    break;
                case "offline":
                    packager = new OfflineInstallerPackager();
                    break;
                default:
                    Util.DisplayWarning("Invalid installer type found.", new InvalidDataException());
                    packager = GetDummyPackager();
                    break;
            }
            packager.SetConfig(config);
            return packager;
        }

        private static AbstractPackager GetDummyPackager()
        {
            return new OnlineInstallerPackager();
        }
    }
}
