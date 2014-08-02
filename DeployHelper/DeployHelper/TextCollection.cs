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
            public const Boolean IsSubDirectoryToDelete = true;
            public const Boolean IsToRemoveAfterUpload = true;
        }

        # endregion
    }
}
