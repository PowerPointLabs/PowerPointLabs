using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using PowerPointLabs.FunctionalTestInterface.Impl;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;

namespace PowerPointLabs.SaveLab
{
    internal class SaveLabConfigSaveFile
    {
        private const string DefaultSaveMasterFolderName = @"\PowerPointLabs Saved Presentation";
        private const string DefaultSaveCategoryName = "My Presentation";
        private const string SaveRootFolderConfigFileName = "SaveRootFolder.config";

        private readonly string _defaultSaveMasterFolderPrefix =
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        private string _configFilePath;

        # region Properties
        public string DefaultCategory { get; set; }
        # endregion

        # region Constructor
        public SaveLabConfigSaveFile(string appDataFolder)
        {
            if (!PowerPointLabsFT.IsFunctionalTestOn)
            {
                SaveLabSettings.SaveFolderPath = _defaultSaveMasterFolderPrefix + DefaultSaveMasterFolderName;
                DefaultCategory = DefaultSaveCategoryName;

                ReadSaveLabConfig(appDataFolder);
            }
            else
            {
                // if it's in FT, use new temp shape root folder every time
                string tmpPath = TempPath.GetTempTestFolder();
                int hash = DateTime.Now.GetHashCode();
                SaveLabSettings.SaveFolderPath = tmpPath + DefaultSaveMasterFolderName + hash;
                DefaultCategory = DefaultSaveCategoryName + hash;
                _configFilePath = tmpPath + "SaveRootFolder" + hash;
            }
        }
        # endregion

        # region Destructor
        ~SaveLabConfigSaveFile()
        {
            // flush shape root folder & default category info to the file
            using (StreamWriter fileWriter = File.CreateText(_configFilePath))
            {
                fileWriter.WriteLine(SaveLabSettings.SaveFolderPath);
                fileWriter.WriteLine(DefaultCategory);
                
                fileWriter.Close();
            }
        }
        # endregion

        # region Helper Functions
        private void ReadSaveLabConfig(string appDataFolder)
        {
            _configFilePath = Path.Combine(appDataFolder, SaveRootFolderConfigFileName);

            if (File.Exists(_configFilePath) &&
                (new FileInfo(_configFilePath)).Length != 0)
            {
                using (StreamReader reader = new StreamReader(_configFilePath))
                {
                    SaveLabSettings.SaveFolderPath = reader.ReadLine();
                    
                    // if we have a default category setting
                    if (reader.Peek() != -1)
                    {
                        DefaultCategory = reader.ReadLine();
                    }

                    reader.Close();
                }
            }
        }
        # endregion
    }
}