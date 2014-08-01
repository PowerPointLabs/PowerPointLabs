using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DeployHelper
{
    class Cleaner
    {
        #region Clean up
        private const Boolean IsSubDirectoryToDelete = true;

        public void CleanUp()
        {
            Directory.Delete(_dirPptLabsZipFolder, IsSubDirectoryToDelete);
            Directory.Delete(DirLocalPathToUpload, IsSubDirectoryToDelete);
            File.Delete(DirPptLabsZipPath);
        }
        #endregion
    }
}
