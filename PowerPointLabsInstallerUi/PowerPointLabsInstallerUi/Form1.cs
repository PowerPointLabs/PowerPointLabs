using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Windows.Forms;

namespace PowerPointLabsInstallerUi
{
    public partial class Form1 : Form
    {
        private readonly string _installerZipAddress = Application.StartupPath + "\\data.zip";
        private const string TextButtonClose = "Close";
        private const string TextButtonRunning = "Running...";
        private const string ErrorWindowTitle = "PowerPointLabs Installer";
        private const string UrlForVstoRuntim = "http://www.comp.nus.edu.sg/~pptlabs/vsto-redirect.html";
        private const string UrlForPptlabsOnlineInstaller = "http://www.comp.nus.edu.sg/~pptlabs/download-78563/PowerPointLabs.zip";

        private readonly string _onlineInstallerZipAddress = Path.Combine(Path.GetTempPath(),
            @"PowerPointLabsInstaller\olInstaller.zip");

        private readonly string _targetInstallFolder;

        public Form1()
        {
            InitializeComponent();

            // handle special char case for EURO user
            _targetInstallFolder = Path.Combine(
                (IsSpecialCharPresentInInstallPath() 
                    ? Path.GetPathRoot(Environment.SystemDirectory) 
                    : Path.GetTempPath()),
                @"PowerPointLabsInstaller");
        }

        /// <summary>
        /// If there are special characters (eg é) present in the install path,
        /// the offline installer (ClickOnce) will fail to install. Thus need to install it to the root path.
        /// </summary>
        /// <returns></returns>
        private bool IsSpecialCharPresentInInstallPath()
        {
            return new Uri(Path.GetTempPath()).AbsolutePath.Replace("/", "\\") != Path.GetTempPath();
        }

        /// <summary>
        /// main behavior
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            if (button1.Text != TextButtonClose)
            {
                //detection for VSTO runtime + its config
                if (!IsVstoRuntimeValid())
                {
                    var dialogResult = MessageBox.Show(
                        "For the PowerPointLabs to work properly, your computer needs to have " +
                        "Visual Studio 2010 Tools for Office (VSTO) Runtime from Microsoft.\n\n" +
                        "Click Yes button to download it, or click No button to continue the installation anyway.",
                        ErrorWindowTitle,
                        MessageBoxButtons.YesNoCancel, MessageBoxIcon.Error);
                    if (dialogResult == DialogResult.Yes)
                    {
                        Process.Start(UrlForVstoRuntim);
                        return;
                    }
                    else if (dialogResult == DialogResult.Cancel)
                    {
                        return;
                    }
                } 
                else if (!IsVstoConfigValid())
                {
                    var vstoConfigDir = GetVstoConfigDir();
                    if (MessageBox.Show(
                        "In order to install our add-in, you need to rename the file [VSTOInstaller.exe.Config] in the folder" +
                        "\n[" + vstoConfigDir + "]\n to the new filename [VSTOInstaller.exe.Config.backup]\n\n" +
                        "After that, click OK button to continue.",
                        ErrorWindowTitle, 
                        MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
                        == DialogResult.Cancel)
                    {
                        return;
                    }
                }
                
                //run installation files
                button1.Enabled = false;
                button1.Text = TextButtonRunning;

                //normal offline installer
                Boolean isUnzipSuccessful = UnzipInstaller(_installerZipAddress);
                if (isUnzipSuccessful)
                {
                    RunInstaller();
                }
                button1.Enabled = true;
                button1.Text = TextButtonClose;
            }
            else
            {
                Close();
            }
        }

        private void AfterOnlineInstallerDownload()
        {
            //unzip online installer
            var isUnzipSuccessful = UnzipInstaller(_onlineInstallerZipAddress);
            //then run it
            if (isUnzipSuccessful)
            {
                MessageBox.Show("In order to install our add-in, please click 'yes' button to allow changes.", 
                    "PowerPointLabs Installer");
                RunInstaller();
            }
            button1.Enabled = true;
            button1.Text = TextButtonClose;
        }

        private void WhenDownloadFailure()
        {
            button1.Enabled = true;
            button1.Text = TextButtonClose;
        }

        private static bool IsVstoRuntimeValid()
        {
            var runtimeExistList = new List<bool>();
            Boolean result = false;

            var directoriesForProgramFiles = GetProgramFilesDirectories();
            foreach(string dir in directoriesForProgramFiles)
            {
                var directoryForVstoRuntime = Path.Combine(dir, @"Common Files\Microsoft Shared\VSTO\10.0");
                runtimeExistList.Add(Directory.Exists(directoryForVstoRuntime));
            }
            //if VSTO runtime folder does not exist --> invalid
            foreach (var isRuntimeExist in runtimeExistList)
            {
                result = result || isRuntimeExist;
            }
            return result;
        }

        private static bool IsVstoConfigValid()
        {
            var configExistList = new List<bool>();
            Boolean result = false;

            var directoriesForProgramFiles = GetProgramFilesDirectories();
            foreach (string dir in directoriesForProgramFiles)
            {
                var directoryForVstoConfig = Path.Combine(dir, 
                    @"Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe.Config");
                configExistList.Add(File.Exists(directoryForVstoConfig));
            }
            foreach (var isConfigExist in configExistList)
            {
                result = result || isConfigExist;
            }
            //if VSTO config file exists --> invalid
            return !result;
        }

        private static string GetVstoConfigDir()
        {
            var directoriesForProgramFiles = GetProgramFilesDirectories();
            foreach (string dir in directoriesForProgramFiles)
            {
                var directoryForVstoConfig = Path.Combine(dir,
                    @"Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe.Config");
                if (File.Exists(directoryForVstoConfig))
                {
                    return directoryForVstoConfig;
                }
            }
            return "";
        }

        private static List<string> GetProgramFilesDirectories()
        {
            var result = new List<string>();
            //For 64-bit Windows
            if (8 == IntPtr.Size
                || (!String.IsNullOrEmpty(Environment.GetEnvironmentVariable("PROCESSOR_ARCHITEW6432"))))
            {
                //C:\Program Files(x86)
                result.Add(Environment.GetEnvironmentVariable("ProgramFiles(x86)"));
                //C:\Program Files
                result.Add(Environment.GetEnvironmentVariable("ProgramW6432"));
            }
            //For 32-bit Windows
            else
            {
                //C:\Program Files
                result.Add(Environment.GetEnvironmentVariable("ProgramFiles"));
            }
            return result;
        } 

        private void RunInstaller()
        {
            try
            {
                var process = new Process
                {
                    StartInfo =
                    {
                        FileName = Path.Combine(_targetInstallFolder, "setup.exe"),
                        WindowStyle = ProcessWindowStyle.Hidden
                    }
                };
                process.Start();
                process.WaitForExit();
            }
            catch (Exception e)
            {
                PowerPointLabs.Views.ErrorDialogWrapper.ShowDialog("Failed to install",
                    "An error occurred while installing PowerPointLabs", e);
            }
        }

        private Boolean UnzipInstaller(String installerZipAddress)
        {
            try
            {
                var installerZip = ZipStorer.Open(installerZipAddress, FileAccess.Read);
                var zipDir = installerZip.ReadCentralDir();
                foreach (var file in zipDir)
                {
                    installerZip.ExtractFile(file,
                        Path.Combine(_targetInstallFolder, file.FilenameInZip));
                }
                installerZip.Close();
                return true;
            }
            catch (Exception e)
            {
                PowerPointLabs.Views.ErrorDialogWrapper.ShowDialog("Failed to install",
                    "An error occurred while installing PowerPointLabs", e);
            }
            return false;
        }
    }
}
