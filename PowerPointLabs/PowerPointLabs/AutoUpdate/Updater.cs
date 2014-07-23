using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Xml;
using DocumentFormat.OpenXml.Drawing;
using Path = System.IO.Path;

namespace PowerPointLabs.AutoUpdate
{
    class Updater
    {
        //TODO: make it configurable
        const string vstoAddress = "http://www.comp.nus.edu.sg/~pptlabs/download_testDeploy/dev/PowerPointLabs.vsto";
        readonly string destVstoAddress = Path.Combine(Path.GetTempPath(), "PowerPointLabs.update.vsto");
        const string offlineInstallerAddress = "http://www.comp.nus.edu.sg/~pptlabs/download_testDeploy/dev/PowerPointLabsOffline.zip";
        readonly string destOfflineInstallerAddress = Path.Combine(Path.GetTempPath(), "PowerPointLabs.zip");

        public void TryUpdate()
        {
            var downloader = new Downloader();
            downloader.Get(vstoAddress, destVstoAddress)
                .After(AfterVstoDownloadHandler)
                .Start();
        }

        private void AfterVstoDownloadHandler()
        {
            if (GetVstoVersion(destVstoAddress) != TextCollection.CurrentVersion)
            {
                var downloader = new Downloader();
                downloader.Get(offlineInstallerAddress, destOfflineInstallerAddress)
                    .After(AfterInstallerDownloadHandler)
                    .Start();
            }
        }

        private void AfterInstallerDownloadHandler()
        {
            UnzipInstaller(destOfflineInstallerAddress);
            RunInstaller();
        }

        private static void RunInstaller()
        {
            try
            {
                var process = new Process
                {
                    StartInfo =
                    {
                        FileName = Path.Combine(Path.GetTempPath(), @"PowerPointLabsInstaller\setup.exe"),
                        WindowStyle = ProcessWindowStyle.Hidden
                    }
                };
                process.Start();
                process.WaitForExit();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "RunInstaller");
            }
        }

        private void UnzipInstaller(String installerZipAddress)
        {
            var installerZip = ZipStorer.Open(installerZipAddress, FileAccess.Read);
            var zipDir = installerZip.ReadCentralDir();
            foreach (var file in zipDir)
            {
                installerZip.ExtractFile(file,
                    Path.Combine(Path.GetTempPath(), @"PowerPointLabsInstaller\" + file.FilenameInZip));
            }
            installerZip.Close();
        }

        private string GetVstoVersion(String vstoDirectory)
        {
            var currentVsto = new XmlDocument();
            try
            {
                currentVsto.Load(vstoDirectory);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetVstoVersion");
            }
            var vstoNode = currentVsto.GetElementsByTagName("assemblyIdentity")[0];

            if (vstoNode.Attributes != null)
            {
                return vstoNode.Attributes["version"].Value;
            }
            else
            {
                return "";
            }
        }
    }
}
