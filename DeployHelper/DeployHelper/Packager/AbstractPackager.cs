using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace DeployHelper.Packager
{
    abstract class AbstractPackager
    {
        protected DeployConfig Config;

        public void SetConfig(DeployConfig config)
        {
            Config = config;
        }

        // Get x86 special folder, ref:
        // http://stackoverflow.com/questions/3540930/getting-syswow64-directory-using-32-bit-application
        [DllImport("shell32.dll")]
        public static extern bool SHGetSpecialFolderPath(IntPtr hwndOwner, [Out]StringBuilder lpszPath, int nFolder, bool fCreate);

        string GetSystemDirectory()
        {
            StringBuilder path = new StringBuilder(260);
            SHGetSpecialFolderPath(IntPtr.Zero, path, 0x0029, false);
            return path.ToString();
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
                    // Only get iexpress from x86 special folder 
                    // (system32 under x86, syswow64 under x64)
                    FileName = Path.Combine(GetSystemDirectory(), "iexpress.exe"),
                    Arguments = "/N " + iexpressSedDirectory,
                    WindowStyle = ProcessWindowStyle.Hidden
                }
            };
            process.Start();
            process.WaitForExit();
        }
    }
}
