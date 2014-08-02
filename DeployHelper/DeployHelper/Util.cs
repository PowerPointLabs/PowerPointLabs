using System;
using System.IO;

namespace DeployHelper
{
    class Util
    {
        #region Helper functions

        public static void ConsoleWriteWithColor(String content, ConsoleColor color)
        {
            Console.ForegroundColor = color;
            Console.Write(content);
            Console.ResetColor();
        }

        public static void PrepareWelcomeInfo()
        {
            Console.WriteLine("Checklist before deploy:");
            Console.Write("1. Have you updated the version number in the ");
            ConsoleWriteWithColor("About ", ConsoleColor.Yellow);
            Console.WriteLine("button?");
            Console.Write("2. Is there newer version of ");
            ConsoleWriteWithColor("Pptlabs tutorial", ConsoleColor.Yellow);
            Console.WriteLine("? If there is," +
                              " you need to update it in the web server.");
            Console.WriteLine("Press any key to continue");
            Console.ReadKey();
        }

        public static void IgnoreException()
        {
        }

        public static void DisplayWarning(string content, Exception e)
        {
            ConsoleWriteWithColor(content, ConsoleColor.Red);
            throw new InvalidOperationException(content, e);
        }

        public static void DisplayDone(string content)
        {
            ConsoleWriteWithColor(content + "\n", ConsoleColor.Green);
        }

        public static string AddQuote(string dir)
        {
            return "\"" + dir + "\"";
        }

        //Taken from http://msdn.microsoft.com/en-us/library/cc148994.aspx
        //How to: Copy, Delete, and Move Files and Folders (C# Programming Guide)
        public static void CopyFolder(string sourcePath, string destPath, bool isOverWritten)
        {
            if (!Directory.Exists(sourcePath)) return;
            var files = Directory.GetFiles(sourcePath);

            // Copy the files
            foreach (var file in files)
            {
                // Use static Path methods to extract only the file name from the path.
                var fileName = Path.GetFileName(file);
                if (fileName != null)
                {
                    var destFile = Path.Combine(destPath, fileName);
                    File.Copy(file, destFile, isOverWritten);
                }
            }
        }

        public static void CreateDirectory(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }

        public static void DisplayEndMessage()
        {
            DisplayDone("All Done!");
            Console.WriteLine("Have a nice day :)");
        }

        #endregion
    }
}
