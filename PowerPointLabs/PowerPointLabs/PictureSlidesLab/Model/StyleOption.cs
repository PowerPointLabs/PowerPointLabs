using System;
using System.ComponentModel;
using System.IO;
using System.Xml.Serialization;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    [Serializable]
    public partial class StyleOption
    {
        public StyleOption()
        {
            Init();
        }

        #region Initialization
        private void Init()
        {
            foreach (PropertyDescriptor property in TypeDescriptor.GetProperties(this))
            {
                var myAttribute = (DefaultValueAttribute) property
                    .Attributes[typeof(DefaultValueAttribute)];
                if (myAttribute != null)
                {
                    property.SetValue(this, myAttribute.Value);
                }
            }
        }
        #endregion

        #region IO serialization
        /// Taken from http://stackoverflow.com/a/14663848

        /// <summary>
        /// Saves to an xml file
        /// </summary>
        /// <param name="filename">File path of the new xml file</param>
        public void Save(string filename)
        {
            try
            {
                using (var writer = new StreamWriter(filename))
                {
                    var serializer = new XmlSerializer(GetType());
                    serializer.Serialize(writer, this);
                    writer.Flush();
                }
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.Log("Failed to save Picture Slides Lab Style Options: " + e.StackTrace, "Error");
            }
        }

        /// <summary>
        /// Load an object from an xml file
        /// </summary>
        /// <param name="filename">Xml file name</param>
        /// <returns>The object created from the xml file</returns>
        public static StyleOption Load(string filename)
        {
            try
            {
                using (var stream = File.OpenRead(filename))
                {
                    var serializer = new XmlSerializer(typeof(StyleOption));
                    var opt = serializer.Deserialize(stream) as StyleOption;
                    return opt ?? new StyleOption();
                }
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.Log("Failed to load Picture Slides Lab Style Options: " + e.StackTrace, "Error");
                return new StyleOption();
            }
        }

        # endregion
        
    }
}
