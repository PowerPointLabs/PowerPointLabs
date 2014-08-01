using System;
using System.IO;

namespace DeployHelper
{
    class ManifestManager
    {
        private ManifestEditor _editor;
        private ManifestSignee _signee;

        public ManifestManager()
        {
            _editor = new ManifestEditor(TextCollection.Config.DirBuildManifest);

            var argsForSignManifest =
                "-sign " + Util.AddQuote(TextCollection.Config.DirBuildManifest) +
                " -certfile " + Util.AddQuote(TextCollection.Config.ConfigDirKey);
            var argsForSignVsto =
                "-update " + Util.AddQuote(TextCollection.Config.DirVsto) +
                " -appmanifest " + Util.AddQuote(TextCollection.Config.DirBuildManifest) +
                " -certfile " + Util.AddQuote(TextCollection.Config.ConfigDirKey);
            _signee = new ManifestSignee(argsForSignManifest, argsForSignVsto, TextCollection.Config.ConfigDirMage);
        }

        public void EditManifest()
        {
            if (!IsPatched())
            {
                VerifyKeyFileExist();
                _editor.ModifyManifest();
                _signee.Sign();
                //overwrite build vsto file with resigned new vsto file
                File.Copy(TextCollection.Config.DirVsto, TextCollection.Config.DirBuildVsto, TextCollection.Const.IsOverWritten);
                Util.DisplayDone(TextCollection.Const.DonePatched);
            }
            else
            {
                Util.DisplayDone(TextCollection.Const.DonePatchedAlready);
            }
        }

        private static void VerifyKeyFileExist()
        {
            if (!File.Exists(TextCollection.Config.ConfigDirKey))
            {
                Util.DisplayWarning(TextCollection.Const.ErrorInvalidKeyOrMageDir, new FileNotFoundException());
            }
        }

        private Boolean IsPatched()
        {
            return _editor.IsPatched();
        }
    }
}
