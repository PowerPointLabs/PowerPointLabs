using System;
using System.IO;

using PowerPointLabs.FunctionalTestInterface.Impl;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ShapesLab
{
    internal class ShapesLabConfigSaveFile
    {
#pragma warning disable 0618
        private const string DefaultShapeMasterFolderName = @"\PowerPointLabs Custom Shapes";
        private const string DefaultShapeCategoryName = "My Shapes";
        private const string ShapeRootFolderConfigFileName = "ShapeRootFolder.config";

        private readonly string _defaultShapeMasterFolderPrefix =
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        private string _configFilePath;

        # region Properties
        public string DefaultCategory { get; set; }
        # endregion

        # region Constructor
        public ShapesLabConfigSaveFile(string appDataFolder)
        {
            if (!PowerPointLabsFT.IsFunctionalTestOn)
            {
                ShapesLabSettings.SaveFolderPath = _defaultShapeMasterFolderPrefix + DefaultShapeMasterFolderName;
                DefaultCategory = DefaultShapeCategoryName;

                ReadShapeLabConfig(appDataFolder);
            }
            else
            {
                // if it's in FT, use new temp shape root folder every time
                string tmpPath = TempPath.GetTempTestFolder();
                int hash = DateTime.Now.GetHashCode();
                ShapesLabSettings.SaveFolderPath = tmpPath + DefaultShapeMasterFolderName + hash;
                DefaultCategory = DefaultShapeCategoryName + hash;
                _configFilePath = tmpPath + "ShapeRootFolder" + hash;
            }
        }
        # endregion

        # region Destructor
        ~ShapesLabConfigSaveFile()
        {
            // flush shape root folder & default category info to the file
            using (StreamWriter fileWriter = File.CreateText(_configFilePath))
            {
                fileWriter.WriteLine(ShapesLabSettings.SaveFolderPath);
                fileWriter.WriteLine(DefaultCategory);
                
                fileWriter.Close();
            }
        }
        # endregion

        # region Helper Functions
        private void ReadShapeLabConfig(string appDataFolder)
        {
            _configFilePath = Path.Combine(appDataFolder, ShapeRootFolderConfigFileName);

            if (File.Exists(_configFilePath) &&
                (new FileInfo(_configFilePath)).Length != 0)
            {
                using (StreamReader reader = new StreamReader(_configFilePath))
                {
                    ShapesLabSettings.SaveFolderPath = reader.ReadLine();
                    
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