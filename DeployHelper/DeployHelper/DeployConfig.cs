using System;
using System.Diagnostics;
using System.IO;

namespace DeployHelper
{
    class DeployConfig
    {
        public static readonly string DirCurrent = Environment.CurrentDirectory;
        public static readonly string DirVsto = DirCurrent + @"\PowerPointLabs.vsto";
        public static readonly string DirConfig = DirCurrent + @"\DeployHelper.conf";
        public static readonly string DirOfflineInstallerFolder = DirCurrent + @"\offline";
        public static readonly string DirOnlineInstallerFolder = DirCurrent + @"\online";
        public static readonly string DirLocalPathToUpload = DirCurrent + @"\PowerPointLabs_upload";
        public static readonly string DirAppFilesToUpload = DirCurrent + @"\PowerPointLabs_upload\Application Files";

        //currentFolder\Application Files\PowerPointLabs_A_B_C_D\
        public string DirBuild;

        //currentFolder\Application Files\PowerPointLabs_A_B_C_D\PowerPointLabs.dll.manifest
        public string DirBuildManifest;

        //currentFolder\Application Files\PowerPointLabs_A_B_C_D\PowerPointLabs.dll.config.deploy
        public string DirBuildConfig;

        //currentFolder\Application Files\PowerPointLabs_A_B_C_D\PowerPointLabs.vsto
        public string DirBuildVsto;

        //PowerPointLabs_A_B_C_D
        public string DirBuildName;

        public string ConfigDirMage;
        public string ConfigDirKey;
        public string ConfigSftpAddress;
        public string ConfigSftpUser;
        public string ConfigDevPath;
        public string ConfigReleasePath;

        public string Version;
        public string VersionMajor;
        public string VersionMinor;
        public string VersionBuild;
        public string VersionRevision;

        public string ReleaseType;
        public string InstallerType;
        public string ConfigVersion;
        public string ReleaseAddress;
        public string DevAddress;

        public void VerifyConfig()
        {
            if (ReleaseType != "release" && ReleaseType != "dev")
            {
                Util.DisplayWarning(TextCollection.Const.ErrorInvalidConfig + " Release type not correct.", new Exception());
            }

            if (InstallerType != "online" && InstallerType != "offline")
            {
                Util.DisplayWarning(TextCollection.Const.ErrorInvalidConfig + " Installer type not correct.", new Exception());
            }

//            if (Version != ConfigVersion)
//            {
//                Util.DisplayWarning(TextCollection.Const.ErrorInvalidConfig + " Version number different.", new Exception());
//            }

            //auto-correct installation folder url
            string argForInstallationFolderUrl = "-url=";
            if (ReleaseType == "release" && InstallerType == "online")
            {
                argForInstallationFolderUrl += ReleaseAddress;
            }
            else if (ReleaseType == "dev" && InstallerType == "online")
            {
                argForInstallationFolderUrl += DevAddress;
            }

            var process = new Process
            {
                StartInfo =
                {
                    FileName = Path.Combine(DirCurrent, "setup.exe"),
                    Arguments = argForInstallationFolderUrl,
                    WindowStyle = ProcessWindowStyle.Hidden
                }
            };
            process.Start();
            process.WaitForExit();
        }
    }
}
