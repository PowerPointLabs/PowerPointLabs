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
            _signee = new ManifestSignee(_config);
        }

        public void EditManifest()
        {
            if (!IsPatched())
            {
                ModifyManifest();
            }
            Sign();
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
                if (!_signee.IsSuccessful())
                {
                    throw new Exception();
                }
                //overwrite build vsto file with resigned new vsto file
                File.Copy(DeployConfig.DirVsto, _config.DirBuildVsto, TextCollection.Const.IsOverWritten);
                Util.DisplayDone(TextCollection.Const.DonePatched);
            }
            catch (Exception e)
            {
                Util.DisplayWarning(TextCollection.Const.ErrorInvalidKeyOrMageDirOrPassword, e);
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
