using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace PowerPointLabs.Utils
{
    class PathUtil
    {
        public static String GetTempTestFolder()
        {
            return Path.Combine(Path.GetTempPath(), "PowerPointLabsTest\\");
        }
    }
}
