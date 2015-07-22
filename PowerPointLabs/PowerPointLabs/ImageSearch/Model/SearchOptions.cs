using System;
using System.IO;
using System.Xml.Serialization;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ImageSearch.Model
{
    [Serializable]
    public class SearchOptions : Notifiable
    {
        private string _searchEngineId;

        public string SearchEngineId
        {
            get
            {
                return _searchEngineId;
            }
            set
            {
                _searchEngineId = value;
                OnPropertyChanged("SearchEngineId");
            }
        }

        private string _apiKey;

        public string ApiKey
        {
            get
            {
                return _apiKey;
            }
            set
            {
                _apiKey = value;
                OnPropertyChanged("ApiKey");
            }
        }

        private int _colorType;

        public int ColorType
        {
            get
            {
                return _colorType;
            }
            set
            {
                _colorType = value;
                OnPropertyChanged("ColorType");
            }
        }

        private int _dominantColor;

        public int DominantColor
        {
            get
            {
                return _dominantColor;
            }
            set
            {
                _dominantColor = value;
                OnPropertyChanged("DominantColor");
            }
        }

        private int _imageType;

        public int ImageType
        {
            get
            {
                return _imageType;
            }
            set
            {
                _imageType = value;
                OnPropertyChanged("ImageType");
            }
        }

        private int _imageSize;

        public int ImageSize
        {
            get { return _imageSize; }
            set
            {
                _imageSize = value;
                OnPropertyChanged("ImageSize");
            }
        }

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
                    SearchEngineId = Common.Base64Encode(SearchEngineId);
                    ApiKey = Common.Base64Encode(ApiKey);
                    serializer.Serialize(writer, this);
                    writer.Flush();
                }
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.Log("Failed to save Images Lab Search Options: " + e.StackTrace, "Error");
            }
        }

        /// <summary>
        /// Load an object from an xml file
        /// </summary>
        /// <param name="filename">Xml file name</param>
        /// <returns>The object created from the xml file</returns>
        public static SearchOptions Load(string filename)
        {
            try
            {
                using (var stream = File.OpenRead(filename))
                {
                    var serializer = new XmlSerializer(typeof(SearchOptions));
                    var opt = serializer.Deserialize(stream) as SearchOptions;
                    if (opt != null)
                    {
                        opt.SearchEngineId = Common.Base64Decode(opt.SearchEngineId);
                        opt.ApiKey = Common.Base64Decode(opt.ApiKey);
                    }
                    else
                    {
                        opt = new SearchOptions();
                        opt.Init();
                    }
                    return opt;
                }
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.Log("Failed to load Images Lab Search Options: " + e.StackTrace, "Error");
                var opt = new SearchOptions();
                opt.Init();
                return opt;
            }
        }

        private void Init()
        {
            SearchEngineId = "";
            ApiKey = "";
            ColorType = 0;
            DominantColor = 0;
            ImageType = 0;
            ImageSize = 2;
        }

        public string GetColorType()
        {
            switch (ColorType)
            {
                case 0:
                    return "color";
                case 1:
                    return "gray";
                case 2:
                    return "mono";
                default:
                    return "";
            }
        }

        public string GetDominantColor()
        {
            switch (DominantColor)
            {
                case 0:
                    return "none";
                case 1:
                    return "black";
                case 2:
                    return "blue";
                case 3:
                    return "brown";
                case 4:
                    return "gray";
                case 5:
                    return "green";
                case 6:
                    return "pink";
                case 7:
                    return "purple";
                case 8:
                    return "teal";
                case 9:
                    return "white";
                case 10:
                    return "yellow";
                default:
                    return "";
            }
        }

        public string GetImageType()
        {
            switch (ImageType)
            {
                case 0:
                    return "photo";
                case 1:
                    return "clipart";
                case 2:
                    return "lineart";
                case 3:
                    return "news";
                case 4:
                    return "face";
                default:
                    return "";
            }
        }

        public string GetImageSize()
        {
            switch (ImageSize)
            {
                case 0:
                    return "huge";
                case 1:
                    return "icon";
                case 2:
                    return "large";
                case 3:
                    return "medium";
                case 4:
                    return "small";
                case 5:
                    return "xlarge";
                case 6:
                    return "xxlarge";
                default:
                    return "";

            }
        }

        # endregion
    }
}
