using System;
using DeployHelper.Packager;

namespace DeployHelper
{
    class InstallerPackager
    {
        #region Produce Zip

        private readonly DeployConfig _config;

        public InstallerPackager(DeployConfig config)
        {
            _config = config;
        }

        public void ProducePackage()
        {
            Console.WriteLine("Zipping...");

            var packager = PackagerFactory.GetPackager(_config);
            packager.Produce();
            
            Util.DisplayDone(TextCollection.Const.DoneZipped);
        }

        #endregion
    }
}
