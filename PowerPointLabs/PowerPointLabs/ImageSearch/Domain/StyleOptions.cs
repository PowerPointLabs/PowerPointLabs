using System;
using System.IO;
using System.Xml.Serialization;
using PowerPointLabs.WPF.Observable;

namespace PowerPointLabs.ImageSearch.Domain
{
    [Serializable]
    public class StyleOptions : Model
    {
        # region UI related prop
        private bool _isUseOriginalTextFormat;

        public bool IsUseOriginalTextFormat
        {
            get
            {
                return _isUseOriginalTextFormat;
            }
            set
            {
                _isUseOriginalTextFormat = value;
                OnPropertyChanged("IsUseOriginalTextFormat");
            }
        }

        private int _fontFamily;

        public int FontFamily
        {
            get { return _fontFamily; }
            set
            {
                _fontFamily = value;
                OnPropertyChanged("FontFamily");
            }
        }

        private int _fontSizeIncrease;

        public int FontSizeIncrease
        {
            get { return _fontSizeIncrease; }
            set
            {
                _fontSizeIncrease = value;
                OnPropertyChanged("FontSizeIncrease");
            }
        }

        private string _fontColor;

        public string FontColor
        {
            get { return _fontColor; }
            set
            {
                _fontColor = value;
                OnPropertyChanged("FontColor");
            }
        }

        private string _overlayColor;

        public string OverlayColor
        {
            get { return _overlayColor; }
            set
            {
                _overlayColor = value;
                OnPropertyChanged("OverlayColor");
            }
        }

        private int _transparency;

        // for overlay
        public int Transparency
        {
            get { return _transparency; }
            set
            {
                _transparency = value;
                OnPropertyChanged("Transparency");
            }
        }
        # endregion

        # region Logic
        public void Init()
        {
            IsUseOriginalTextFormat = false;
            FontFamily = 1;
            FontSizeIncrease = 10;
            FontColor = "#FFFFFF";
            OverlayColor = "#000000";
            Transparency = 85;
        }

        public string GetFontFamily()
        {
            switch (FontFamily)
            {
                case 0:
                    return "Segoe UI";
                case 1:
                    return "Segoe UI Light";
                case 2:
                    return "Calibri";
                case 3:
                    return "Calibri Light";
                case 4:
                    return "";
                default:
                    return "Segoe UI";
            }
        }
        # endregion

        # region IO serialization
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
                PowerPointLabsGlobals.Log("Failed to save Images Lab Style Options: " + e.StackTrace, "Error");
            }
        }

        /// <summary>
        /// Load an object from an xml file
        /// </summary>
        /// <param name="filename">Xml file name</param>
        /// <returns>The object created from the xml file</returns>
        public static StyleOptions Load(string filename)
        {
            try
            {
                using (var stream = File.OpenRead(filename))
                {
                    var serializer = new XmlSerializer(typeof (StyleOptions));
                    var opt = serializer.Deserialize(stream) as StyleOptions;
                    return opt ?? CreateDefault();
                }
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.Log("Failed to load Images Lab Style Options: " + e.StackTrace, "Error");
                return CreateDefault();
            }
        }

        private static StyleOptions CreateDefault()
        {
            var opt = new StyleOptions();
            opt.Init();
            return opt;
        }

        # endregion
    }
}
