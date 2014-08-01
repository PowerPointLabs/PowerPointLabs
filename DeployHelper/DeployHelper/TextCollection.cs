using System;

namespace DeployHelper
{
    class TextCollection
    {
        #region Const

        public class Const
        {
            public const string ErrorNoConfig = "Can't Find Config.";
            public const string ErrorNoVsto = "Can't Find VSTO.";
            public const string ErrorNoManifest = "Can't Find Manifest For This Version.";
            public const string ErrorInvalidKeyOrMageDir = "Invalid Mage or Key Directory.";
            public const string ErrorZipFilesMissing = "Some files to zip are missing.";
            public const string ErrorNetworkFailed = "Can't Connect The Server.";

            public const string DonePatched = "Patched.";
            public const string DonePatchedAlready = "Patched already.";
            public const string DoneZipped = "Zipped.";
            public const string DoneSftpConnected = "SFTP Connected.";
            public const string DoneUploaded = "Uploaded.";

            public const string InfoEnterPassword = "Enter SFTP password: ";
            public const string InfoFileUploading = "Uploading files";
            public const string InfoChooseVersion = "Which version to upload? [dev|release]: ";

            public const string VarDev = "dev";
            public const string VarRelease = "release";

            public const Boolean IsOverWritten = true;
        }

        #endregion

        # region To be init
        public class Config
        {
            public static readonly string DirCurrent = Environment.CurrentDirectory;
            public static readonly string DirVsto = DirCurrent + @"\PowerPointLabs.vsto";
            public static readonly string DirConfig = DirCurrent + @"\DeployHelper.conf";

            //currentFolder\Application Files\PowerPointLabs_A_B_C_D\
            public static string DirBuild;

            //currentFolder\Application Files\PowerPointLabs_A_B_C_D\PowerPointLabs.dll.manifest
            public static string DirBuildManifest;

            //currentFolder\Application Files\PowerPointLabs_A_B_C_D\PowerPointLabs.dll.config.deploy
            public static string DirBuildConfig;

            //currentFolder\Application Files\PowerPointLabs_A_B_C_D\PowerPointLabs.vsto
            public static string DirBuildVsto;

            //PowerPointLabs_A_B_C_D
            public static string DirBuildName;

            public static string ConfigDirMage;
            public static string ConfigDirKey;
            public static string ConfigSftpAddress;
            public static string ConfigSftpUser;
            public static string ConfigDevPath;
            public static string ConfigReleasePath;

            public static string Version;
            public static string VersionMajor;
            public static string VersionMinor;
            public static string VersionBuild;
            public static string VersionRevision;

            public static string ReleaseType;
            public static string InstallerType;
        }

        # endregion
    }
}
