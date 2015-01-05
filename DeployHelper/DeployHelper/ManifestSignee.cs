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

        public ManifestSignee(DeployConfig config)
        {
            Console.Write(TextCollection.Const.InfoEnterCertificatePassword);
            var password = Console.ReadLine();
            var argsForSignManifest =
                "-sign " + Util.AddQuote(config.DirBuildManifest) +
                " -certfile " + Util.AddQuote(config.ConfigDirKey) +
                " -pwd " + Util.AddQuote(password);
            var argsForSignVsto =
                "-update " + Util.AddQuote(DeployConfig.DirVsto) +
                " -appmanifest " + Util.AddQuote(config.DirBuildManifest) +
                " -certfile " + Util.AddQuote(config.ConfigDirKey) +
                " -pwd " + Util.AddQuote(password);
            _argsForSignManifest = argsForSignManifest;
            _argsForSignVsto = argsForSignVsto;
            _mageDirectory = config.ConfigDirMage;
        }

        public void Sign()
        {
            SignManifest();
            SignVsto();
        }

        private void SignVsto()
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

        private void SignManifest()
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

        #endregion
    }
}
