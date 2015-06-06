using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace FunctionalTest.util
{
    class PathUtil
    {
        public static String GetParentFolder(String path)
        {
            return Directory.GetParent(path).FullName;
        }

        public static String GetParentFolder(String path, int loopCount)
        {
            if (loopCount <= 0) 
                return path;

            String parPath = GetParentFolder(path);
            return GetParentFolder(parPath, --loopCount);
        }

        public static String GetTempPath(String fileName)
        {
            return Path.GetTempPath() + fileName;
        }
    }
}
