using System;
using System.IO;

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

        public static string GetDocTestPath()
        {
            //To get the location the assembly normally resides on disk or the install directory
            var path = new Uri(
                Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase))
                .LocalPath;
            var parPath = PathUtil.GetParentFolder(path, 4);
            return Path.Combine(parPath, "doc\\test\\");
        }

        public static string GetTestFailurePath()
        {
            return GetDocTestPath() + "TestFailed\\";
        }

        public static string GetTestFailurePresentationPath(string presentationName)
        {
            return GetTestFailurePath() + presentationName;
        }

        public static string GetDocTestPresentationPath(string presentationName)
        {
            return GetDocTestPath() + presentationName;
        }
    }
}
