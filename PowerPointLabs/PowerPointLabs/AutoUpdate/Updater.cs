using System;
using System.IO;
using System.IO.Compression;
using System.Xml;
using PowerPointLabs.ActionFramework.Common.Log;
using Path = System.IO.Path;

namespace PowerPointLabs.AutoUpdate
{
    class Updater
    {
        private readonly string _vstoAddress;
        private readonly string _offlineInstallerAddress;
        private readonly string _destVstoAddress = Path.Combine(Path.GetTempPath(), TextCollection.VstoName);
        private readonly string _destOfflineInstallerAddress = Path.Combine(Path.GetTempPath(), TextCollection.InstallerName);
        private readonly string _targetInstallFolder;

        public Updater()
        {
            //init files address
            switch (Properties.Settings.Default.ReleaseType)
            {
                case "dev":
                    _vstoAddress = Properties.Settings.Default.DevAddr + TextCollection.VstoName;
                    _offlineInstallerAddress = Properties.Settings.Default.DevAddr + TextCollection.InstallerName;
                    break;
                case "release":
                    _vstoAddress = Properties.Settings.Default.ReleaseAddr + TextCollection.VstoName;
                    _offlineInstallerAddress = Properties.Settings.Default.ReleaseAddr + TextCollection.InstallerName;
                    break;
                default:
                    _vstoAddress = "";
                    _offlineInstallerAddress = "";
                    break;
            }

            // handle special char case for EURO user
            _targetInstallFolder = Path.Combine(
                (IsSpecialCharPresentInInstallPath()
                    ? Path.GetPathRoot(Environment.SystemDirectory)
                    : Path.GetTempPath()),
                @"PowerPointLabsInstaller");
        }

        public void TryUpdate()
        {
            if (IsInstallerTypeOnline())
            {
                return;
            }

            new Downloader()
                .Get(_vstoAddress, _destVstoAddress)
                .After(AfterVstoDownloadHandler)
                .Start();
        }

        private void AfterInstallerDownloadHandler()
        {
            Unzip(_destOfflineInstallerAddress);
            //No need to run it, ppt will auto exec it when run next time
        }

        private void AfterVstoDownloadHandler()
        {
            var version = GetVstoVersion(_destVstoAddress);
            if (IsTheNewestVersion(version))
            {
                return;
            }

            new Downloader()
                .Get(_offlineInstallerAddress, _destOfflineInstallerAddress)
                .After(AfterInstallerDownloadHandler)
                .Start();
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
                Logger.LogException(e, "GetVstoVersion");
            }
            var vstoNode = currentVsto.GetElementsByTagName("assemblyIdentity")[0];

            return vstoNode.Attributes != null
                ? vstoNode.Attributes["version"].Value
                : "";
        }

        private bool IsInstallerTypeOnline()
        {
            return Properties.Settings.Default.InstallerType.ToLower() != "offline"
                   || _vstoAddress == "" 
                   || _offlineInstallerAddress == "";
        }

        /// <summary>
        /// If there are special characters (eg é) present in the install path,
        /// the offline installer (ClickOnce) will fail to install. Thus need to install it to the root path.
        /// </summary>
        /// <returns></returns>
        private bool IsSpecialCharPresentInInstallPath()
        {
            return new Uri(Path.GetTempPath()).AbsolutePath.Replace("/", "\\") != Path.GetTempPath();
        }

        private static bool IsTheNewestVersion(string version)
        {
            return version != "" && version == Properties.Settings.Default.Version;
        }

        private void Unzip(String installerZipAddress)
        {
            var installerZip = ZipStorer.Open(installerZipAddress, FileAccess.Read);
            var zipDir = installerZip.ReadCentralDir();
            foreach (var file in zipDir)
            {
                installerZip.ExtractFile(file,
                    Path.Combine(_targetInstallFolder, file.FilenameInZip));
            }
            installerZip.Close();
        }
    }
}
