using System;
using System.Diagnostics;
using System.IO;
using System.Xml;
using WinSCP;
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
//  setup.bat, PowerPointLabs Quick Tutorial.pptx, and folder data.
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
            PrepareWelcomeInfo();
            try
            {
                //Reference on What It Does
                //http://msdn.microsoft.com/en-us/library/vstudio/dd465291(v=vs.100).aspx
                //Walkthrough: Copying a Document to the End User Computer after a ClickOnce Installation
                ReadConfig();
                ModifyManifest();
                ReSign();

                ProduceZip();
                SftpUpload();
                CleanUp();
                DisplayEndMessage();
            }
            catch
            {
                IgnoreException();
            }
            finally
            {
                Console.ReadKey();
            }
        }

        #region Const
        private const string ErrorNoConfig = "Can't Find Config.";
        private const string ErrorNoVsto = "Can't Find VSTO.";
        private const string ErrorNoManifest = "Can't Find Manifest For This Version.";
        private const string ErrorInvalidKeyOrMageDir = "Invalid Mage or Key Directory.";
        private const string ErrorZipFilesMissing = "Some files to zip are missing. Are data folder," + 
            " readme.txt, setup.bat, and quick tutorial inside your publish folder?";
        private const string ErrorNetworkFailed = "Can't Connect The Server.";

        private const string DonePatched = "Patched.";
        private const string DonePatchedAlready = "Patched already.";
        private const string DoneZipped = "Zipped.";
        private const string DoneSftpConnected = "SFTP Connected.";
        private const string DoneUploaded = "Uploaded.";

        private const string InfoEnterPassword = "Enter SFTP password: ";
        private const string InfoFileUploading = "Uploading files...";
        private const string InfoChooseVersion = "Which version to upload? [dev|release]: ";

        private const string VarDev = "dev";
        private const string VarRelease = "release";

        private const Boolean IsOverWritten = true;
        #endregion
        #region Helper functions

        private static void ConsoleWriteWithColor(String content, ConsoleColor color)
        {
            Console.ForegroundColor = color;
            Console.Write(content);
            Console.ResetColor();
        }

        private static void PrepareWelcomeInfo()
        {
            Console.WriteLine("Checklist before deploy:");
            Console.Write("1. Have you updated the version number in the ");
            ConsoleWriteWithColor("About ", ConsoleColor.Yellow);
            Console.WriteLine("button?");
            Console.Write("2. Is there newer version of ");
            ConsoleWriteWithColor("Pptlabs tutorial", ConsoleColor.Yellow);
            Console.WriteLine("? If there is," +
                              " you need to update it in the Pptlabs project and in the zip file as well");
            Console.WriteLine("Press any key to continue");
            Console.ReadKey();
        }

        private static void IgnoreException()
        {
        }

        private static void DisplayWarning(string content)
        {
            ConsoleWriteWithColor(content, ConsoleColor.Red);
            throw new InvalidOperationException(content);
        }

        private static void DisplayDone(string content)
        {
            ConsoleWriteWithColor(content + "\n", ConsoleColor.Green);
        }

        private static string AddQuote(string dir)
        {
            return "\"" + dir + "\"";
        }

        //Taken from http://msdn.microsoft.com/en-us/library/cc148994.aspx
        //How to: Copy, Delete, and Move Files and Folders (C# Programming Guide)
        private static void CopyFolder(string sourcePath, string destPath, bool isOverWritten)
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
        #endregion
        #region Read Config

        private static readonly string DirCurrent = Environment.CurrentDirectory;
        private static readonly string DirVsto = DirCurrent + @"\PowerPointLabs.vsto";

        private static string _dirBuild;        //currentFolder\Application Files\PowerPointLabs_A_B_C_D\
        private static string _dirManifest;     //currentFolder\Application Files\PowerPointLabs_A_B_C_D\PowerPointLabs.dll.manifest
        private static string _dirDestVsto;     //currentFolder\Application Files\PowerPointLabs_A_B_C_D\PowerPointLabs.vsto
        private static string _dirNameNewestVer;//PowerPointLabs_A_B_C_D

        private static string _configDirMage;
        private static string _configDirKey;
        private static string _configSftpAddress;
        private static string _configSftpUser;
        private static string _configDevPath;
        private static string _configReleasePath;
        private static string[] _configContent;

        private static string _version;
        private static string _versionMajor;
        private static string _versionMinor;
        private static string _versionBuild;
        private static string _versionRevision;
        private static XmlDocument _currentVsto;

        private static void ReadConfig()
        {
            InitConfig();
            InitVersion();
            InitDir();
        }

        private static void InitConfig()
        {
            var configDirectory = DirCurrent + @"\DeployHelper.conf";
            LoadConfigContent(configDirectory);
            //index here refers to the line number in DeployHelper.conf
            //TODO: any better way to config?
            _configDirMage = _configContent[1];
            _configDirKey = _configContent[3];
            _configSftpAddress = _configContent[5];
            _configSftpUser = _configContent[7];
            _configDevPath = _configContent[9];
            _configReleasePath = _configContent[11];
        }

        private static void LoadConfigContent(string configDirectory)
        {
            _configContent = new string[] {};
            try
            {
                _configContent = File.ReadAllLines(configDirectory);
            }
            catch
            {
                DisplayWarning(ErrorNoConfig);
            }
        }

        private static void InitVersion()
        {
            LoadCurrentVsto();
            var vstoNode = _currentVsto.GetElementsByTagName("assemblyIdentity")[0];
            Debug.Assert(vstoNode.Attributes != null);

            if (vstoNode.Attributes != null) _version = vstoNode.Attributes["version"].Value;
            //Assume that version follows this style: Major.Minor.Build.Revision
            var versionDetails = _version.Split('.');

            _versionMajor = versionDetails[0];
            _versionMinor = versionDetails[1];
            _versionBuild = versionDetails[2];
            _versionRevision = versionDetails[3];
        }

        private static void LoadCurrentVsto()
        {
            _currentVsto = new XmlDocument();
            try
            {
                _currentVsto.Load(DirVsto);
            }
            catch
            {
                DisplayWarning(ErrorNoVsto);
            }
        }

        private static void InitDir()
        {
            _dirNameNewestVer = "PowerPointLabs_"
                                + _versionMajor + "_" + _versionMinor + "_" + _versionBuild + "_" + _versionRevision;
            _dirBuild = DirCurrent + @"\Application Files\" + _dirNameNewestVer;
            _dirManifest = _dirBuild + @"\PowerPointLabs.dll.manifest";
            _dirDestVsto = _dirBuild + @"\PowerPointLabs.vsto";
        }
        #endregion
        #region Modify Manifest

        private static Boolean _isPatched;
        private static XmlDocument _manifest;
        private static XmlDocument _manifestBackup;

        private static void ModifyManifest()
        {
            Console.WriteLine("Patching...");
            LoadManifest();
            if (!IsPatched())
            {
                PatchManifest();
            }
            else
            {
                DisplayDone(DonePatchedAlready);
            }
        }

        private static void LoadManifest()
        {
            _manifest = new XmlDocument();
            _manifestBackup = new XmlDocument();
            try
            {
                _manifest.Load(_dirManifest);
                _manifestBackup.Load(_dirManifest);
            }
            catch
            {
                DisplayWarning(ErrorNoManifest);
            }
        }

        private static Boolean IsPatched()
        {
            return _isPatched = _manifest.GetElementsByTagName("vstav3:postAction").Count != 0;
        }

        private static void PatchManifest()
        {
            //Patch Content for Manifest
            //************************************************************
            //<vstav3:postActions>
            //  <vstav3:postAction>
            //      <vstav3:entryPoint class="PowerPointLabs.PostInstall">
            //          <assemblyIdentity 
            //          name="PostInstall" 
            //          version="{$version}" 
            //          language="neutral" 
            //          processorArchitecture="msil"/>
            //      </vstav3:entryPoint>
            //      <vstav3:postActionData>
            //      </vstav3:postActionData>
            //  </vstav3:postAction>
            //</vstav3:postActions>
            //************************************************************
            const string vstaNamespaceUri = "urn:schemas-microsoft-com:vsta.v3";

            //setup nodes
            XmlNode addInNode = _manifest.GetElementsByTagName("vstav3:addIn")[0];
            XmlNode updateNode = _manifest.GetElementsByTagName("vstav3:update")[0];
            XmlElement postActionsNode = _manifest.CreateElement("vstav3", "postActions",
                vstaNamespaceUri);
            XmlElement postActionNode = _manifest.CreateElement("vstav3", "postAction",
                vstaNamespaceUri);
            XmlElement entryPointNode = _manifest.CreateElement("vstav3", "entryPoint",
                vstaNamespaceUri);
            entryPointNode.SetAttribute("class", "PowerPointLabs.PostInstall");
            XmlElement postActionDataNode = _manifest.CreateElement("vstav3", "postActionData",
                vstaNamespaceUri);

            //insert and append them
            addInNode.InsertAfter(postActionsNode, updateNode);
            postActionsNode.AppendChild(postActionNode);
            postActionNode.AppendChild(entryPointNode);
            entryPointNode.InnerXml = "<assemblyIdentity " +
                                      "name=" + AddQuote("PostInstall") + " " +
                                      "version=" + AddQuote(_version) + " " +
                                      "language=" + AddQuote("neutral") + " " +
                                      "processorArchitecture=" + AddQuote("msil") + "/>";
            postActionNode.AppendChild(postActionDataNode);
            _manifest.Save(_dirManifest);
        }
        #endregion
        #region Re-Sign

        private static string _argsForSignManifest;
        private static string _argsForSignVsto;

        private static void ReSign()
        {
            if (_isPatched) return;
            CheckDirKeyExist();
            InitReSignArgs();
            SignManifest();
            SignVsto();
            //overwrite old vsto file with resigned new vsto file
            File.Copy(DirVsto, _dirDestVsto, IsOverWritten);
            DisplayDone(DonePatched);
        }

        private static void CheckDirKeyExist()
        {
            if (!File.Exists(_configDirKey))
            {
                HandleReSignFailure();
            }
        }

        private static void InitReSignArgs()
        {
            _argsForSignManifest =
                "-sign " + AddQuote(_dirManifest) + " -certfile " + AddQuote(_configDirKey);
            _argsForSignVsto =
                "-update " + AddQuote(DirVsto) + " -appmanifest " + AddQuote(_dirManifest) +
                " -certfile " + AddQuote(_configDirKey);
        }

        private static void SignVsto()
        {
            try
            {
                var process = new Process
                {
                    StartInfo =
                    {
                        FileName = _configDirMage,
                        Arguments = _argsForSignVsto,
                        WindowStyle = ProcessWindowStyle.Hidden
                    }
                };
                process.Start();
                process.WaitForExit();
            }
            catch
            {
                HandleReSignFailure();
            }
        }

        private static void SignManifest()
        {
            try
            {
                var process = new Process
                {
                    StartInfo =
                    {
                        FileName = _configDirMage,
                        Arguments = _argsForSignManifest,
                        WindowStyle = ProcessWindowStyle.Hidden
                    }
                };
                process.Start();
                process.WaitForExit();
            }
            catch
            {
                HandleReSignFailure();
            }
        }

        private static void HandleReSignFailure()
        {
            //Restore manifest file
            _manifestBackup.Save(_dirManifest);
            DisplayWarning(ErrorInvalidKeyOrMageDir);
        }

        #endregion
        #region Produce Zip

        private static readonly string DirPptLabsZipPath = DirCurrent + @"\PowerPointLabs.zip";
        private static readonly string DirPptLabsZipFolder = DirCurrent + @"\PowerPointLabs";

        private static void ProduceZip()
        {
            Console.WriteLine("Zipping...");
            SetupBinExe();
            SetupZipFolder();
            CreateZipFile();
            DisplayDone(DoneZipped);
        }

        private static void SetupZipFolder()
        {
            //copy folder data, ReadMe.txt, PowerPointLabs Quick Tutorial.pptx, and setup.bat
            //into the powerPointLabsZipFolder; zip them together to produce PowerPointLabs.zip
            //source:
            var dirDataFolder = DirCurrent + @"\data\";
            var dirReadMe = DirCurrent + @"\ReadMe.txt";
            var dirQuickTutorial = DirCurrent + @"\PowerPointLabs Quick Tutorial.pptx";
            var dirSetupBat = DirCurrent + @"\setup.bat";
            //dest:
            var dirZipDataFolder = DirPptLabsZipFolder + @"\data\";
            var dirZipReadMe = DirPptLabsZipFolder + @"\ReadMe.txt";
            var dirZipQuickTutorial = DirPptLabsZipFolder + @"\PowerPointLabs Quick Tutorial.pptx";
            var dirZipSetupBat = DirPptLabsZipFolder + @"\setup.bat";

            //if dir not set up yet, then need to create them first
            CreateDirectory(DirPptLabsZipFolder);
            CreateDirectory(dirZipDataFolder);

            try
            {
                CopyFolder(dirDataFolder, dirZipDataFolder, IsOverWritten);
                File.Copy(dirReadMe, dirZipReadMe, IsOverWritten);
                File.Copy(dirQuickTutorial, dirZipQuickTutorial, IsOverWritten);
                File.Copy(dirSetupBat, dirZipSetupBat, IsOverWritten);
            }
            catch
            {
                DisplayWarning(ErrorZipFilesMissing);
            }
        }

        private static void SetupBinExe()
        {
            //rename setup.exe to bin.exe, and copy it into the data folder
            var setupExeDirectory = DirCurrent + @"\setup.exe";
            var destSetupExeDirectory = DirCurrent + @"\data\bin.exe";
            File.Copy(setupExeDirectory, destSetupExeDirectory, IsOverWritten);
        }

        private static void CreateZipFile()
        {
            //remove the old zip file, if any
            if (File.Exists(DirPptLabsZipPath))
            {
                File.Delete(DirPptLabsZipPath);
            }
            System.IO.Compression.ZipFile.CreateFromDirectory(DirPptLabsZipFolder, DirPptLabsZipPath);
        }

        private static void CreateDirectory(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }
        #endregion
        #region SFTP upload

        private static readonly string DirLocalPathToUpload = DirCurrent + @"\PowerPointLabs_upload";
        private static readonly string DirAppFilesInLocalPath = DirLocalPathToUpload + @"\Application Files";
        private static readonly string ZipPathToUpload = DirLocalPathToUpload + @"\PowerPointLabs.zip";
        private static readonly string VstoPathToUpload = DirLocalPathToUpload + @"\PowerPointLabs.vsto";

        private const Boolean IsToRemoveAfterUpload = true;

        private static Boolean _isReleased;

        private static void SftpUpload()
        {
            try
            {
                var sessionOptions = SetupSessionOptions();
                using (var session = new Session())
                {
                    //create a folder to upload and put PowerPointLabs.zip, PowerPointLabs.vsto,
                    //and Application Files/PowerPointLabs_X_X_X_X (the newest ver) into the folder
                    var dirNewestVerFolder = DirAppFilesInLocalPath + "\\" + _dirNameNewestVer;

                    CreateDirectory(DirLocalPathToUpload);
                    CreateDirectory(DirAppFilesInLocalPath);
                    CreateDirectory(dirNewestVerFolder);

                    Console.WriteLine("Connecting the server...");
                    session.Open(sessionOptions);
                    if (session.Opened)
                    {
                        DisplayDone(DoneSftpConnected);
                        Console.WriteLine(InfoFileUploading);

                        var remotePath = SetupRemotePath();
                        var transferOptions = SetupTransferOptions();
                        ConstructRemoteFolder(session, remotePath, transferOptions);

                        // Copy files into DirLocalPathToUpload
                        CopyFolder(_dirBuild, dirNewestVerFolder, IsOverWritten);
                        File.Copy(DirPptLabsZipPath, ZipPathToUpload, IsOverWritten);
                        File.Copy(DirVsto, VstoPathToUpload, IsOverWritten);

                        Console.WriteLine("Uploading...");
                        UploadLocalFile(session, remotePath, transferOptions);
                        DisplayDone(DoneUploaded);
                    }
                    else
                    {
                        DisplayWarning(ErrorNetworkFailed);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error during SFTP uploading: {0}", e);
                CleanUp();
                DisplayWarning(ErrorNetworkFailed);
            }
        }

        private static void UploadLocalFile(Session session, string remotePath, TransferOptions transferOptions)
        {
            var transferResult = session.PutFiles(
                DirLocalPathToUpload + @"\*",
                remotePath,
                !IsToRemoveAfterUpload,
                transferOptions);

            transferResult.Check();
        }

        private static void ConstructRemoteFolder(Session session, string remotePath, TransferOptions transferOptions)
        {
            // Construct folder with permissions first
            try
            {
                session.PutFiles(
                    DirLocalPathToUpload + @"\*",
                    remotePath,
                    !IsToRemoveAfterUpload,
                    transferOptions);
            }
            catch (InvalidOperationException)
            {
                if (session.Opened)
                {
                    IgnoreException();
                }
                else
                {
                    throw;
                }
            }
        }

        private static TransferOptions SetupTransferOptions()
        {
            var transferOptions = new TransferOptions {TransferMode = TransferMode.Binary};
            var permissions = new FilePermissions {Octal = "644"};
            transferOptions.FilePermissions = permissions;
            return transferOptions;
        }

        private static string SetupRemotePath()
        {
            Console.Write(InfoChooseVersion);
            var versionToUpload = Console.ReadLine();
            string remotePath;
            switch (versionToUpload)
            {
                case VarDev:
                    remotePath = _configDevPath;
                    break;
                case VarRelease:
                    remotePath = _configReleasePath;
                    _isReleased = true;
                    break;
                default:
                    remotePath = _configDevPath;
                    break;
            }
            return remotePath;
        }

        //TODO: 1. hide server and username info, how? let user type once?
        private static SessionOptions SetupSessionOptions()
        {
            Console.Write(InfoEnterPassword);
            var password = Console.ReadLine();
            while (password == null || password.Trim() == "")
            {
                Console.Write(InfoEnterPassword);
                password = Console.ReadLine();
            }

            var sessionOptions = new SessionOptions
            {
                Protocol = Protocol.Sftp,
                HostName = _configSftpAddress,
                UserName = _configSftpUser,
                Password = password,
                PortNumber = 22, //TODO: make it configurable
                GiveUpSecurityAndAcceptAnySshHostKey = true
            };
            return sessionOptions;
        }

        #endregion
        #region Clean up

        private const Boolean IsSubDirectoryToDelete = true;

        private static void CleanUp()
        {
            Directory.Delete(DirPptLabsZipFolder, IsSubDirectoryToDelete);
            Directory.Delete(DirLocalPathToUpload, IsSubDirectoryToDelete);
            File.Delete(DirPptLabsZipPath);
        }

        private static void DisplayEndMessage()
        {
            DisplayDone("All Done!");
            if (_isReleased)
            {
                PrepareEndMessage();
            }
            Console.WriteLine("Have a nice day :)");
        }

        private static void PrepareEndMessage()
        {
            Console.Write("Remember to merge from dev branch into ");
            ConsoleWriteWithColor("release ", ConsoleColor.Yellow);
            Console.WriteLine("branch.");
        }
        #endregion
    }
}
