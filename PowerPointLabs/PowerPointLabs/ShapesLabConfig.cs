using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace PowerPointLabs
{
    internal class ShapesLabConfig
    {
        private const string DefaultShapeMasterFolderName = @"\PowerPointLabs Custom Shapes";
        private const string DefaultShapeCategoryName = "My Shapes";
        private const string ShapeRootFolderConfigFileName = "ShapeRootFolder.config";

        private readonly string _defaultShapeMasterFolderPrefix =
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        private string _configFilePath;

        # region Properties
        public string ShapeRootFolder { get; set; }
        public string DefaultCategory { get; set; }
        # endregion

        # region Constructor
        public ShapesLabConfig(string appDataFolder)
        {
            ShapeRootFolder = _defaultShapeMasterFolderPrefix + DefaultShapeMasterFolderName;
            DefaultCategory = DefaultShapeCategoryName;

            ReadShapeLabConfig(appDataFolder);
        }
        # endregion

        # region Destructor
        ~ShapesLabConfig()
        {
            // flush shape root folder & default category info to the file
            using (var fileWriter = File.CreateText(_configFilePath))
            {
                fileWriter.WriteLine(ShapeRootFolder);
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
                using (var reader = new StreamReader(_configFilePath))
                {
                    ShapeRootFolder = reader.ReadLine();
                    
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
