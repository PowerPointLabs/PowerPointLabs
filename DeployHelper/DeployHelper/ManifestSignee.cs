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
        private Boolean _isSuccess = true;

        public ManifestSignee(DeployConfig config)
        {
            Console.Write(TextCollection.Const.InfoEnterCertificatePassword);
            var password = Console.ReadLine();
            var argsForSignManifest =
                "-sign " + Util.AddQuote(config.DirBuildManifest) +
                " -certfile " + Util.AddQuote(config.ConfigDirKey);
            var argsForSignVsto =
                "-update " + Util.AddQuote(DeployConfig.DirVsto) +
                " -appmanifest " + Util.AddQuote(config.DirBuildManifest) +
                " -certfile " + Util.AddQuote(config.ConfigDirKey);
            if (password != null && password.Trim() != "")
            {
                argsForSignManifest += " -pwd " + Util.AddQuote(password);
                argsForSignVsto += " -pwd " + Util.AddQuote(password);
            }
            _argsForSignManifest = argsForSignManifest;
            _argsForSignVsto = argsForSignVsto;
            _mageDirectory = config.ConfigDirMage;
        }

        public void Sign()
        {
            SignManifest();
            SignVsto();
        }

        public bool IsSuccessful()
        {
            return _isSuccess;
        }

        private void SignVsto()
        {
            var process = new Process
            {
                StartInfo =
                {
                    FileName = _mageDirectory,
                    Arguments = _argsForSignVsto,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    RedirectStandardOutput = true,
                    UseShellExecute = false
                }
            };
            process.Start();
            if (!process.StandardOutput.ReadToEnd().Contains("success"))
            {
                _isSuccess = false;
            }
            process.WaitForExit();
        }

        private void SignManifest()
        {
            var process = new Process
            {
                StartInfo =
                {
                    FileName = _mageDirectory,
                    Arguments = _argsForSignManifest,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    RedirectStandardOutput = true,
                    UseShellExecute = false
                }
            };
            process.Start();
            if (!process.StandardOutput.ReadToEnd().Contains("success"))
            {
                _isSuccess = false;
            }
            process.WaitForExit();
        }

        #endregion
    }
}
