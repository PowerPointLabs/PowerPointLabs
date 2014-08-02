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
            var iexpressSedDirectory = "";
            var fileNames = Directory.EnumerateFiles(folderDirectory);
            foreach (var fileName in fileNames)
            {
                if (!fileName.ToLower().EndsWith("sed"))
                    continue;

                iexpressSedDirectory = fileName;
                break;
            }

            if (iexpressSedDirectory.Trim() == "")
            {
                throw new InvalidDataException("Cannot find SED file to run IExpress");
            }

            //command: iexpress /N "SED_Directory"
            var process = new Process
            {
                StartInfo =
                {
                    FileName = "iexpress",
                    Arguments = "/N " + iexpressSedDirectory + "",
                    WindowStyle = ProcessWindowStyle.Hidden
                }
            };
            process.Start();
            process.WaitForExit();
        }
    }
}
