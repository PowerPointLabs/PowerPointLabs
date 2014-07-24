using System;
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

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (button1.Text != TextButtonClose)
            {
                button1.Enabled = false;
                button1.Text = TextButtonRunning;

                
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

        private static void RunInstaller()
        {
            try
            {
                var process = new Process
                {
                    StartInfo =
                    {
                        FileName = Path.Combine(Path.GetTempPath(), @"PowerPointLabsInstaller\setup.exe"),
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
                        Path.Combine(Path.GetTempPath(), @"PowerPointLabsInstaller\" + file.FilenameInZip));
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
