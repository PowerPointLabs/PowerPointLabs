using System;
using System.IO;
using System.Linq;
using System.Security.Permissions;

namespace PowerPointLabs.Utils
{
    public static class FileDir
    {
        private const string FolderThumbnailFile = "Thumbs.db";

        # region Folder Operations
        public static bool CopyFolder(string oldPath, string newPath)
        {
            bool copySuccess = true;

            // create subfolder during recursions
            if (!Directory.Exists(newPath))
            {
                Directory.CreateDirectory(newPath);
            }

            // copy files in a folder first
            string[] files = Directory.GetFiles(oldPath);

            foreach (string file in files)
            {
                string name = Path.GetFileName(file);

                // ignore thumb.db
                if (name == null ||
                    name == FolderThumbnailFile)
                {
                    continue;
                }

                string dest = Path.Combine(newPath, name);

                try
                {
                    FileAttributes fileAttribute = File.GetAttributes(file);
                    
                    try
                    {
                        File.SetAttributes(file, FileAttributes.Normal);
                    }
                    catch (Exception)
                    {
                    }
                    
                    CopyFile(file, dest, fileAttribute);
                }
                catch (Exception)
                {
                    copySuccess = false;
                }
            }

            // then recursively copy contents in subfolders
            string[] folders = Directory.GetDirectories(oldPath);

            foreach (string folder in folders)
            {
                string name = Path.GetFileName(folder);

                if (name == null)
                {
                    continue;
                }

                string dest = Path.Combine(newPath, name);

                copySuccess = copySuccess && CopyFolder(folder, dest);
            }

            return copySuccess;
        }

        public static bool DeleteFolder(string path)
        {
            bool deleteSuccess = true;
            // copy files in a folder first
            string[] files = Directory.GetFiles(path);

            foreach (string file in files)
            {
                string name = Path.GetFileName(file);

                if (name == null)
                {
                    continue;
                }

                try
                {
                    DeleteFile(file);
                }
                catch (Exception)
                {
                    deleteSuccess = false;
                }
            }

            string[] folders = Directory.GetDirectories(path);

            foreach (string folder in folders)
            {
                string name = Path.GetFileName(folder);

                if (name == null)
                {
                    continue;
                }

                deleteSuccess = deleteSuccess && DeleteFolder(folder);
            }

            if (deleteSuccess)
            {
                try
                {
                    Directory.Delete(path);
                }
                catch (Exception)
                {
                    deleteSuccess = false;
                }
            }

            return deleteSuccess;
        }

        public static bool IsDirectoryEmpty(string path)
        {
            return !Directory.EnumerateFileSystemEntries(path).Any();
        }

        public static bool MoveFolder(string oldPath, string newPath)
        {
            return CopyFolder(oldPath, newPath) && DeleteFolder(oldPath);
        }

        public static string GetTemporaryPngFilePath()
        {
            string tempDirectory = Path.GetTempPath();
            Directory.CreateDirectory(tempDirectory);

            string randomFileName = "pptlabs_" + Path.GetRandomFileName();
            string randomFilePath = Path.Combine(tempDirectory, randomFileName);
            return Path.ChangeExtension(randomFilePath, "png");
        }

        /// <summary>
        /// This function sets attribute of all files in a folder and its sub-folder 
        /// to normal.
        /// </summary>
        /// <param name="path">The folder's location.</param>
        public static void NormalizeFolder(string path)
        {
            // copy files in a folder first
            string[] files = Directory.GetFiles(path);

            foreach (string file in files)
            {
                File.SetAttributes(file, FileAttributes.Normal);
            }

            // then recursively copy contents in subfolders
            string[] folders = Directory.GetDirectories(path);

            foreach (string folder in folders)
            {
                NormalizeFolder(folder);
            }
        }
        # endregion

        # region File Operations
        /// <summary>
        /// This function is an integration of copy-without-attribute and copy-with-attribute.
        /// If fileAttribute is not set explicitly, the file is copying without attribute, else
        /// the specified attribute will be set after the file has been copied.
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <param name="destPath"></param>
        /// <param name="fileAttribute"></param>
        public static void CopyFile(string sourcePath, string destPath,
                                        FileAttributes fileAttribute = FileAttributes.Normal)
        {
            if (!File.Exists(sourcePath))
            {
                return;
            }

            File.Copy(sourcePath, destPath);
            File.SetAttributes(destPath, fileAttribute);
        }

        public static void DeleteFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return;
            }

            File.SetAttributes(filePath, FileAttributes.Normal);
            File.Delete(filePath);
        }

        /// <summary>
        /// Attempts to delete the file and does not delete the file if undeletable
        /// </summary>
        public static void TryDeleteFile(string filePath)
        {
            try
            {
                DeleteFile(filePath);
            }
            catch (Exception)
            {
            }
        }

        public static bool IsFileReadable(string filePath)
        {
            FileIOPermission fileIOPermission = new FileIOPermission(FileIOPermissionAccess.Read,
                                            System.Security.AccessControl.AccessControlActions.View,
                                            filePath);

            return fileIOPermission.AllFiles == FileIOPermissionAccess.Read;
        }
        # endregion
    }
}
