using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PowerPointLabs.Views;

namespace PowerPointLabs.Utils
{
    public static class FileAndDirTask
    {
        public static bool CopyFolder(string oldPath, string newPath)
        {
            var copySuccess = true;

            // create subfolder during recursions
            if (!Directory.Exists(newPath))
            {
                Directory.CreateDirectory(newPath);
            }

            // copy files in a folder first
            var files = Directory.GetFiles(oldPath);

            foreach (var file in files)
            {
                var name = Path.GetFileName(file);

                // ignore thumb.db
                if (name == null ||
                    name == "thumb.db") continue;

                var dest = Path.Combine(newPath, name);

                try
                {
                    var fileAttribute = File.GetAttributes(file);
                    File.SetAttributes(file, FileAttributes.Normal);
                    File.Copy(file, dest);
                    File.SetAttributes(dest, fileAttribute);
                }
                catch (Exception)
                {
                    copySuccess = false;
                }
            }

            // then recursively copy contents in subfolders
            var folders = Directory.GetDirectories(oldPath);

            foreach (var folder in folders)
            {
                var name = Path.GetFileName(folder);

                if (name == null) continue;

                var dest = Path.Combine(newPath, name);

                copySuccess = copySuccess && CopyFolder(folder, dest);
            }

            return copySuccess;
        }

        public static void FileDeleteWithAttribute(string filePath,
                                                   FileAttributes fileAttributes = FileAttributes.Normal)
        {
            if (!File.Exists(filePath)) return;

            try
            {
                File.SetAttributes(filePath, fileAttributes);
                File.Delete(filePath);
            }
            catch (Exception e)
            {
                ErrorDialogWrapper.ShowDialog(TextCollection.AccessTempFolderErrorMsg, string.Empty, e);
            }
        }

        public static void FileCopyWithAttribute(string sourcePath, string destPath,
                                                 FileAttributes fileAttributes = FileAttributes.Normal)
        {
            if (!File.Exists(sourcePath)) return;

            // copy the file to temp folder and rename to zip
            try
            {
                File.Copy(sourcePath, destPath);
                File.SetAttributes(destPath, fileAttributes);
            }
            catch (Exception e)
            {
                ErrorDialogWrapper.ShowDialog(TextCollection.AccessTempFolderErrorMsg, string.Empty, e);
            }
        }

        public static bool IsDirectoryEmpty(string path)
        {
            return !Directory.EnumerateFileSystemEntries(path).Any();
        }

        public static bool MoveFolder(string oldPath, string newPath)
        {
            if (!CopyFolder(oldPath, newPath))
            {
                return false;
            }

            try
            {
                NormalizeFolder(oldPath);
                Directory.Delete(oldPath, true);
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        public static void NormalizeFolder(string path)
        {
            // copy files in a folder first
            var files = Directory.GetFiles(path);

            foreach (var file in files)
            {
                File.SetAttributes(file, FileAttributes.Normal);
            }

            // then recursively copy contents in subfolders
            var folders = Directory.GetDirectories(path);

            foreach (var folder in folders)
            {
                NormalizeFolder(folder);
            }
        }
    }
}
