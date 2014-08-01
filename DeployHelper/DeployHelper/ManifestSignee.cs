using System;
using System.Diagnostics;

namespace DeployHelper
{
    class ManifestSignee
    {
        #region Re-Sign

        private readonly string _argsForSignManifest;
        private readonly string _argsForSignVsto;
        private readonly string _mageDirectory;

        public ManifestSignee(String argsForSignManifest, String argsForSignVsto, String mageDirectory)
        {
            _argsForSignManifest = argsForSignManifest;
            _argsForSignVsto = argsForSignVsto;
            _mageDirectory = mageDirectory;
        }

        public void Sign()
        {
            SignManifest();
            SignVsto();
        }

        private void SignVsto()
        {
            try
            {
                var process = new Process
                {
                    StartInfo =
                    {
                        FileName = _mageDirectory,
                        Arguments = _argsForSignVsto,
                        WindowStyle = ProcessWindowStyle.Hidden
                    }
                };
                process.Start();
                process.WaitForExit();
            }
            catch (Exception e)
            {
                Util.DisplayWarning(TextCollection.Const.ErrorInvalidKeyOrMageDir, e);
            }
        }

        private void SignManifest()
        {
            try
            {
                var process = new Process
                {
                    StartInfo =
                    {
                        FileName = _mageDirectory,
                        Arguments = _argsForSignManifest,
                        WindowStyle = ProcessWindowStyle.Hidden
                    }
                };
                process.Start();
                process.WaitForExit();
            }
            catch (Exception e)
            {
                Util.DisplayWarning(TextCollection.Const.ErrorInvalidKeyOrMageDir, e);
            }
        }

        #endregion
    }
}
