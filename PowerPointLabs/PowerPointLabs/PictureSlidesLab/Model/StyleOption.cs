using System;
using System.ComponentModel;
using System.IO;
using System.Xml.Serialization;

using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    /// <summary>
    /// StyleOption provides settings or options for a style.
    /// 
    /// To support any new properties, create a new partial class of 
    /// StyleOption under folder `StyleOption.Partial`.
    /// </summary>
    [Serializable]
    public partial class StyleOption
    {
        public StyleOption()
        {
            Init();
        }

        #region IO serialization

        // TODO move these to StorageUtil

        /// Taken from http://stackoverflow.com/a/14663848

        /// <summary>
        /// Saves to an xml file
        /// </summary>
        /// <param name="filename">File path of the new xml file</param>
        public void Save(string filename)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(filename))
                {
                    XmlSerializer serializer = new XmlSerializer(GetType());
                    serializer.Serialize(writer, this);
                    writer.Flush();
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Failed to save Picture Slides Lab Style Options: " + e.StackTrace);
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
                using (FileStream stream = File.OpenRead(filename))
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(StyleOption));
                    StyleOption opt = serializer.Deserialize(stream) as StyleOption;
                    return opt ?? new StyleOption();
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Failed to load Picture Slides Lab Style Options: " + e.StackTrace);
                return new StyleOption();
            }
        }

        # endregion

        #region Initialization
        private void Init()
        {
            foreach (PropertyDescriptor property in TypeDescriptor.GetProperties(this))
            {
                DefaultValueAttribute myAttribute = (DefaultValueAttribute) property
                    .Attributes[typeof(DefaultValueAttribute)];
                if (myAttribute != null)
                {
                    property.SetValue(this, myAttribute.Value);
                }
            }
        }
        #endregion     
    }
}
