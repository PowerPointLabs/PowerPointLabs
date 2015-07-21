using System;

namespace PowerPointLabs.ImageSearch.Util
{
    class StoragePath
    {
        public static string GetPath(string name)
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\" + name + ".pptlabsconfig";
        }
    }
}
