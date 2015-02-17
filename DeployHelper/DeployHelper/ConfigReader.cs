using System;
using System.IO;
using System.Xml;

namespace DeployHelper
{
    class ConfigReader
    {
        #region Read Config

        private readonly string _currentDirectory;
        private readonly string _configDirectory;
        private readonly string _vstoDirectory;

        private static string _dirBuild;        //currentFolder\Application Files\PowerPointLabs_A_B_C_D\
        private static string _dirBuildManifest;//currentFolder\Application Files\PowerPointLabs_A_B_C_D\PowerPointLabs.dll.manifest
        private static string _dirBuildVsto;    //currentFolder\Application Files\PowerPointLabs_A_B_C_D\PowerPointLabs.vsto
        private static string _dirBuildName;    //PowerPointLabs_A_B_C_D
        private static string _dirBuildConfig;

        private static string _configDirMage;
        private static string _configDirKey;
        private static string _configSftpAddress;
        private static string _configSftpUser;
        private static string _configDevPath;
        private static string _configReleasePath;

        private static string _version;
        private static string _versionMajor;
        private static string _versionMinor;
        private static string _versionBuild;
        private static string _versionRevision;

        private static string _releaseType;
        private static string _installerType;
        private static string _configVersion;
        private static string _releaseAddress;
        private static string _devAddress;

        public ConfigReader()
        {
            _currentDirectory = DeployConfig.DirCurrent;
            _configDirectory = DeployConfig.DirConfig;
            _vstoDirectory = DeployConfig.DirVsto;
        }

        public ConfigReader(string currentDirectory, string configDirectory, string vstoDirectory)
        {
            _currentDirectory = currentDirectory;
            _configDirectory = configDirectory;
            _vstoDirectory = vstoDirectory;
        }

        public ConfigReader ReadConfig()
        {
            //Sequence matters
            InitConfigVariables();
            InitVersionVariables();
            InitDirBuildAddresses();
            InitDeployInfo();
            return this;
        }

        public override string ToString()
        {
            return
                "ConfigDirMage: " + _configDirMage + "\r\n" +
                "ConfigDirKey: " + _configDirKey + "\r\n" +
                "ConfigSftpAddress: " + _configSftpAddress + "\r\n" +
                "ConfigSftpUser: " + _configSftpUser + "\r\n" +
                "ConfigDevPath: " + _configDevPath + "\r\n" +
                "ConfigReleasePath: " + _configReleasePath + "\r\n" +

                "Version: " + _version + "\r\n" +
                "VersionMajor: " + _versionMajor + "\r\n" +
                "VersionMinor: " + _versionMinor + "\r\n" +
                "VersionBuild: " + _versionBuild + "\r\n" +
                "VersionRevision: " + _versionRevision + "\r\n" +

                "DirBuildName: " + _dirBuildName + "\r\n" +
                "DirBuild: " + _dirBuild + "\r\n" +
                "DirBuildManifest: " + _dirBuildManifest + "\r\n" +
                "DirBuildVsto: " + _dirBuildVsto + "\r\n" +
                "DirBuildConfig: " + _dirBuildConfig + "\r\n" +

                "ReleaseType: " + _releaseType + "\r\n" +
                "InstallerType: " + _installerType + "\r\n";
        }

        public DeployConfig ToDeployConfig()
        {
            var config = new DeployConfig
            {
                ConfigDirMage = _configDirMage,
                ConfigDirKey = _configDirKey,
                ConfigSftpAddress = _configSftpAddress,
                ConfigSftpUser = _configSftpUser,
                ConfigDevPath = _configDevPath,
                ConfigReleasePath = _configReleasePath,

                Version = _version,
                VersionMajor = _versionMajor,
                VersionMinor = _versionMinor,
                VersionBuild = _versionBuild,
                VersionRevision = _versionRevision,

                DirBuild = _dirBuild,
                DirBuildName = _dirBuildName,
                DirBuildManifest = _dirBuildManifest,
                DirBuildVsto = _dirBuildVsto,
                DirBuildConfig = _dirBuildConfig,

                ReleaseType = _releaseType.ToLower(),
                InstallerType = _installerType.ToLower(),
                ConfigVersion = _configVersion,
                ReleaseAddress = _releaseAddress,
                DevAddress = _devAddress
            };
            return config;
        }

        private void InitConfigVariables()
        {
            string[] configContent = {};
            try
            {
                configContent = File.ReadAllLines(_configDirectory);
            }
            catch (Exception e)
            {
                Util.DisplayWarning(TextCollection.Const.ErrorNoConfig, e);
            }

            //index here refers to the line number in DeployHelper.conf
            _configDirMage = configContent[1];
            _configDirKey = configContent[3];
            _configSftpAddress = configContent[5];
            _configSftpUser = configContent[7];
            _configDevPath = configContent[9];
            _configReleasePath = configContent[11];
        }

        private void InitVersionVariables()
        {
            var currentVsto = new XmlDocument();
            try
            {
                currentVsto.Load(_vstoDirectory);
            }
            catch (Exception e)
            {
                Util.DisplayWarning(TextCollection.Const.ErrorNoVsto, e);
            }

            var vstoNode = currentVsto.GetElementsByTagName("assemblyIdentity")[0];
            if (vstoNode.Attributes != null)
            {
                _version = vstoNode.Attributes["version"].Value;
            }
            //Assume that version follows this style: Major.Minor.Build.Revision
            var versionDetails = _version.Split('.');
            _versionMajor = versionDetails[0];
            _versionMinor = versionDetails[1];
            _versionBuild = versionDetails[2];
            _versionRevision = versionDetails[3];
        }

        private void InitDirBuildAddresses()
        {
            _dirBuildName = "PowerPointLabs_" +
                _versionMajor + "_" +
                _versionMinor + "_" +
                _versionBuild + "_" +
                _versionRevision;
            _dirBuild = _currentDirectory + @"\Application Files\" +
                _dirBuildName;
            _dirBuildManifest = _dirBuild + @"\PowerPointLabs.dll.manifest";
            _dirBuildVsto = _dirBuild + @"\PowerPointLabs.vsto";
            _dirBuildConfig = _dirBuild + @"\PowerPointLabs.dll.config.deploy";
        }

        private void InitDeployInfo()
        {
            PrintInfo("You are going to deploy PowerPointLabs\r\n" +
                      "version: ", _version);
            try
            {
                var currentConfig = new XmlDocument();
                currentConfig.Load(_dirBuildConfig);

                var vstoNode = currentConfig.GetElementsByTagName("value");
                _releaseType = vstoNode[0].InnerText;
                _installerType = vstoNode[1].InnerText;
                _configVersion = vstoNode[2].InnerText;
                _releaseAddress = vstoNode[4].InnerText;
                _devAddress = vstoNode[5].InnerText;

                PrintInfo("Release Type: ", _releaseType);
                PrintInfo("Installer Type: ", _installerType);
            }
            catch
            {
                Console.WriteLine(TextCollection.Const.ErrorNoConfig);
            }
        }

        private void PrintInfo(string text, string highlightedText)
        {
            Console.Write(text);
            Util.ConsoleWriteWithColor(highlightedText, ConsoleColor.Yellow);
            Console.WriteLine("");
        }
        #endregion
    }
}
