using System;
#region DeployHelper Description
//
//  DeployHelper Class
//  ------------------
//  Simply double click the .exe file to patch PowerPointLabs, 
//  so that it supports PostInstall event (e.g. open tutorial file after install pptlabs),
//  to produce zip file for downloading, and to upload the files onto the PowerPointLabs server.
//
//  HOW TO USE
//
//  For the first time use, you need to setup the followings:
//
//  0. Compile DeployHelper using Visual Studio. .NET 4.5 is required. The output program is under bin/debug or bin/release folder.
//
//  1. Fill in DeployHelper.conf
//  - Mage is a component provided by Visual Studio
//  - Key is inside PowerPointLabs project
//  - SFTP address is the server to upload to
//  - SFTP username is the username used to login the server
//  - Dev path is the installation folder path on the server for dev version PowerPointLabs
//  - Release path is the installation folder path on the server for release version PowerPointLabs
//
//  2. Publish PowerPointlabs; inside the publish folder, it should have setup.exe, PowerPointLabs.vsto, and folder Application Files
//  
//  3. Copy DeployHelper.exe, DeployHelper.conf, WinSCP.exe and WinSCPnet.dll from the output folder to the publish folder
//
//  4. Copy PowerPointLabs.zip to the publish folder and extract it HERE; make sure the publish folder contains ReadMe.txt, 
//  setup.bat, and folder data.
//
//  5. Run DeployHelper.exe and follow the instructions.
//
//  For the next time
//
//  0. Publish PowerPointlabs.
//
//  1. Run DeployHelper.exe and follow the instructions.
//
//  Have a nice day :)
//
//TODO: add testing
#endregion
namespace DeployHelper
{
    class Program
    {
        public static void Main(string[] args)
        {
            Util.PrepareWelcomeInfo();
            try
            {
                # region Init

                var currentDirectory = TextCollection.Config.DirCurrent;
                var configDirectory = TextCollection.Config.DirConfig;
                var vstoDirectory = TextCollection.Config.DirVsto;
                new ConfigReader(currentDirectory, configDirectory, vstoDirectory)
                    .ReadConfig()
                    .WriteToTextCollection();

                # endregion

                # region modify manifest

                //Reference on What It Does: make ClickOnce support PostInstall functionality
                //http://msdn.microsoft.com/en-us/library/vstudio/dd465291(v=vs.100).aspx
                new ManifestManager().EditManifest();

                # endregion

                new InstallerPackager(TextCollection.Config.InstallerType)
                    .ProducePackage();
                new DeployUploader(TextCollection.Config.ReleaseType, TextCollection.Config.InstallerType)
                    .SftpUpload();
                new Cleaner().CleanUp();
                Util.DisplayEndMessage();
            }
            catch
            {
                Util.IgnoreException();
            }
            finally
            {
                Console.ReadKey();
            }
        }
    }
}
