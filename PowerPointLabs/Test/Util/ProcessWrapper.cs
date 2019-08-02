using System.Diagnostics;

namespace Test.Util
{
    /// <summary>
    /// Process wrapper for PowerPoint
    /// </summary>
    public class PPTProcessWrapper
    {
        private const string OpenFilePrefix = "/O";
        public readonly string exePath;
        public readonly string fileName;
        public readonly string workingDirectory;

        private string Arguments
        {
            get
            {
                return OpenFilePrefix + " " + workingDirectory + "\\" + fileName;
            }
        }

        public PPTProcessWrapper(string exePath, string fileName, string workingDirectory = "")
        {
            this.exePath = exePath;
            this.fileName = fileName;
            this.workingDirectory = workingDirectory;
        }

        public Process Start()
        {
            return Process.Start(exePath, Arguments);
        }
    }
}
