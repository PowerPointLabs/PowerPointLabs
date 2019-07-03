using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;

using Presentation = Microsoft.Office.Interop.PowerPoint.Presentation;

namespace PowerPointLabs.ActionFramework.NarrationsLab
{
    public class TempStorage
    {
        private const string tempFolderNamePrefix = @"\PowerPointLabs Temp\";
        private const string SlideXmlSearchPattern = @"slide(\d+)\.xml";
        private const string fileSuffix = ".pptx";
        private const string tempZipName = "tempZip.zip";

        private Presentation _currentPresentation;
        private string _tempPath;

        public string TempPath
        {
            get
            {
                InitStorageIfPresentationChanged();
                if (_tempPath == null)
                {
                    throw new NullReferenceException("TempStorage has not been initialized! Call Setup before using.");
                }
                return _tempPath;
            }
        }

        private void InitStorageIfPresentationChanged()
        {
            Presentation pres = Globals.ThisAddIn.Application.ActivePresentation;
            if (pres == _currentPresentation)
            {
                return;
            }
            _currentPresentation = pres;
            string tempPath = GetPresentationTempFolder(pres.Name);
            if (tempPath == null)
            {
                Logger.Log("TempStorage failed to initialize");
                return;
            }
            // if temp folder doesn't exist, create
            RecreateDirectory(tempPath);
            PrepareMediaFiles(pres, tempPath);

            _tempPath = tempPath;
        }

        private static string GetPresentationTempFolder(string presName)
        {
            string tempName = presName.GetHashCode().ToString(CultureInfo.InvariantCulture);
            string tempPath = Path.GetTempPath() + tempFolderNamePrefix + tempName + @"\";

            return tempPath;
        }

        private static void PrepareMediaFiles(Presentation pres, string tempPath)
        {
            string presFullName = pres.FullName;
            string presName = pres.Name;
            string zipFullPath = tempPath + tempZipName;

            // in case of embedded slides, we need to regulate the file name and full name
            RegulatePresentationName(pres, tempPath, ref presName, ref presFullName);
            if (IsEmptyFile(presFullName))
            {
                return;
            }

            try
            {
                // before we do everything, check if there's an undelete old zip file
                // due to some error
                OverwriteFileUsing(presFullName, zipFullPath);
                ExtractMediaFiles(zipFullPath, tempPath);
            }
            catch (Exception e)
            {
                ErrorDialogBox.ShowDialog(CommonText.ErrorPrepareMedia, "Files cannot be linked.", e);
            }
        }

        private static void OverwriteFileUsing(string presFullName, string zipFullPath)
        {
            try
            {
                FileDir.DeleteFile(zipFullPath);
                FileDir.CopyFile(presFullName, zipFullPath);
            }
            catch (Exception e)
            {
                ErrorDialogBox.ShowDialog(CommonText.ErrorAccessTempFolder, string.Empty, e);
            }
        }

        private static void RecreateDirectory(string tempPath)
        {
            try
            {
                if (Directory.Exists(tempPath))
                {
                    Directory.Delete(tempPath, true);
                }
            }
            catch (Exception e)
            {
                ErrorDialogBox.ShowDialog(CommonText.ErrorCreateTempFolder, string.Empty, e);
            }
            finally
            {
                Directory.CreateDirectory(tempPath);
            }
        }

        private static void RegulatePresentationName(Presentation pres, string tempPath, ref string presName,
            ref string presFullName)
        {
            // this function is used to handle "embed on other application" issue. In this case,
            // all of presentation name, path and full name do not match the usual rule: name is
            // "Untitled", path is empty string and full name is "slide in XX application". We need
            // to regulate these fields properly.

            presName = presName.AppendIfAbsent(fileSuffix);
            if (tempPath == null)
            {
                return;
            }

            // every time when recorder pane is open,
            // save this presentation's copy, which will be used
            // to load audio files later
            pres.SaveCopyAs(tempPath + presName);
            presFullName = tempPath + presName;
        }

        private static void ExtractMediaFiles(string zipFullPath, string tempPath)
        {
            try
            {
                ZipStorer zip = ZipStorer.Open(zipFullPath, FileAccess.Read);
                List<ZipStorer.ZipFileEntry> dir = zip.ReadCentralDir();

                Regex regex = new Regex(SlideXmlSearchPattern);

                foreach (ZipStorer.ZipFileEntry entry in dir)
                {
                    string name = Path.GetFileName(entry.FilenameInZip);

                    if (name?.Contains(".wav") ?? false ||
                        regex.IsMatch(name))
                    {
                        zip.ExtractFile(entry, tempPath + name);
                    }
                }

                zip.Close();

                FileDir.DeleteFile(zipFullPath);
            }
            catch (Exception e)
            {
                ErrorDialogBox.ShowDialog(CommonText.ErrorExtract, "Archived files cannot be retrieved.", e);
            }
        }

        private static bool IsEmptyFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return false;
            }

            FileInfo fileInfo = new FileInfo(filePath);

            return fileInfo.Length == 0;
        }
    }
}
