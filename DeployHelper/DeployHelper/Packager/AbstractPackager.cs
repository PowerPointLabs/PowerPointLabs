using System.Diagnostics;
using System.IO;

namespace DeployHelper.Packager
{
    abstract class AbstractPackager
    {
        protected DeployConfig Config;

        public void SetConfig(DeployConfig config)
        {
            Config = config;
        }

        public abstract void Produce();

        protected void RunIExpress(string folderDirectory)
        {
            //use provided SED file, package the data bundle using iexpress.exe
            var iexpressSedDirectory = folderDirectory + @"\PowerPointLabsOffline.SED";
            if (!File.Exists(iexpressSedDirectory))
            {
                iexpressSedDirectory = folderDirectory + @"\PowerPointLabsOnline.SED";
            }

            //command: iexpress /N "SED_Directory"
            var process = new Process
            {
                StartInfo =
                {
                    FileName = "iexpress",
                    Arguments = "/N " + iexpressSedDirectory,
                    WindowStyle = ProcessWindowStyle.Hidden
                }
            };
            process.Start();
            process.WaitForExit();
        }
    }
}
