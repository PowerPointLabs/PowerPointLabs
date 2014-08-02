using System;
using System.IO;

namespace DeployHelper
{
    class ManifestManager
    {
        private readonly DeployConfig _config;
        private ManifestEditor _editor;
        private ManifestSignee _signee;

        public ManifestManager(DeployConfig config)
        {
            _config = config;
            _editor = new ManifestEditor(_config.DirBuildManifest);

            var argsForSignManifest =
                "-sign " + Util.AddQuote(_config.DirBuildManifest) +
                " -certfile " + Util.AddQuote(_config.ConfigDirKey);
            var argsForSignVsto =
                "-update " + Util.AddQuote(DeployConfig.DirVsto) +
                " -appmanifest " + Util.AddQuote(_config.DirBuildManifest) +
                " -certfile " + Util.AddQuote(_config.ConfigDirKey);
            _signee = new ManifestSignee(argsForSignManifest, argsForSignVsto, _config.ConfigDirMage);
        }

        public void EditManifest()
        {
            if (!IsPatched())
            {
                ModifyManifest();
                Sign();
            }
            else
            {
                Util.DisplayDone(TextCollection.Const.DonePatchedAlready);
            }
        }

        private void ModifyManifest()
        {
            try
            {
                _editor.ModifyManifest();
            }
            catch (Exception e)
            {
                Util.DisplayWarning(TextCollection.Const.ErrorNoManifest, e);
            }
        }

        private void Sign()
        {
            try
            {
                VerifyKeyFileExist();
                _signee.Sign();
                //overwrite build vsto file with resigned new vsto file
                File.Copy(DeployConfig.DirVsto, _config.DirBuildVsto, TextCollection.Const.IsOverWritten);
                Util.DisplayDone(TextCollection.Const.DonePatched);
            }
            catch (Exception e)
            {
                _editor.RestoreManifest();
                Util.DisplayWarning(TextCollection.Const.ErrorInvalidKeyOrMageDir, e);
            }
        }

        private void VerifyKeyFileExist()
        {
            if (!File.Exists(_config.ConfigDirKey))
            {
                throw new FileNotFoundException();
            }
        }

        private Boolean IsPatched()
        {
            return _editor.IsPatched();
        }
    }
}
