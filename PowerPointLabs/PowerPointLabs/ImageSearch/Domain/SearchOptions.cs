using System;
using System.IO;
using System.Xml.Serialization;
using PowerPointLabs.Utils;
using PowerPointLabs.WPF.Observable;

namespace PowerPointLabs.ImageSearch.Domain
{
    [Serializable]
    public class SearchOptions : Model
    {
        # region UI related prop
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

        private int _fileType;

        public int FileType
        {
            get
            {
                return _fileType;
            }
            set
            {
                _fileType = value;
                OnPropertyChanged("FileType");
            }
        }

        // the above options are for google engine
        // below for bing engine..

        private int _searchEngine;

        public int SearchEngine
        {
            get
            {
                return _searchEngine;
            }
            set
            {
                _searchEngine = value;
                OnPropertyChanged("SearchEngine");
            }
        }

        private string _bingApiKey;

        public string BingApiKey
        {
            get
            {
                return _bingApiKey;
            }
            set
            {
                _bingApiKey = value;
                OnPropertyChanged("BingApiKey");
            }
        }

        private int _bingImageSize;

        public int BingImageSize
        {
            get
            {
                return _bingImageSize;
            }
            set
            {
                _bingImageSize = value;
                OnPropertyChanged("BingImageSize");
            }
        }

        private int _bingImageColor;

        public int BingImageColor
        {
            get
            {
                return _bingImageColor;
            }
            set
            {
                _bingImageColor = value;
                OnPropertyChanged("BingImageColor");
            }
        }

        private int _bingImageStyle;

        public int BingImageStyle
        {
            get
            {
                return _bingImageStyle;
            }
            set
            {
                _bingImageStyle = value;
                OnPropertyChanged("BingImageStyle");
            }
        }

        private int _bingImageFace;

        public int BingImageFace
        {
            get
            {
                return _bingImageFace;
            }
            set
            {
                _bingImageFace = value;
                OnPropertyChanged("BingImageFace");
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
                    Encrypt();
                    var serializer = new XmlSerializer(GetType());
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
                        Decrypt(opt);
                    }
                    else
                    {
                        opt = CreateDefault();
                    }
                    return opt;
                }
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.Log("Failed to load Images Lab Search Options: " + e.StackTrace, "Error");
                return CreateDefault();
            }
        }

        private static SearchOptions CreateDefault()
        {
            var opt = new SearchOptions();
            opt.Init();
            return opt;
        }

        private void Encrypt()
        {
            SearchEngineId = Common.Base64Encode(SearchEngineId);
            ApiKey = Common.Base64Encode(ApiKey);
            BingApiKey = Common.Base64Encode(BingApiKey);
        }

        private static void Decrypt(SearchOptions opt)
        {
            opt.SearchEngineId = Common.Base64Decode(opt.SearchEngineId);
            opt.ApiKey = Common.Base64Decode(opt.ApiKey);
            opt.BingApiKey = Common.Base64Decode(opt.BingApiKey);
        }

        # endregion
        
        # region Logic

        private void Init()
        {
            SearchEngineId = "";
            ApiKey = "";
            BingApiKey = "";
            ColorType = 0;
            DominantColor = 0;
            ImageType = 0;
            ImageSize = 2;
            FileType = 0;

            SearchEngine = 0;
            BingImageSize = 2;
            BingImageColor = 0;
            BingImageStyle = 0;
            BingImageFace = 0;
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

        public string GetFileType()
        {
            switch (FileType)
            {
                case 0:
                    return "none";
                case 1:
                    return "png";
                case 2:
                    return "jpg";
                case 3:
                    return "bmp";
                case 4:
                    return "gif";
                default:
                    return "none";
            }
        }

        public string GetBingImageFilters()
        {
            return "Size:" + GetBingImageSize() +
                   "+Color:" + GetBingImageColor() +
                   "+Style:" + GetBingImageStyle() +
                   (StringUtil.IsEmpty(GetBingImageFace()) ? "" : "+Face:" + GetBingImageFace());
        }

        private string GetBingImageSize()
        {
            switch (BingImageSize)
            {
                case 0:
                    return "Small";
                case 1:
                    return "Medium";
                // case 2:
                default:
                    return "Large";
            }
        }

        private string GetBingImageColor()
        {
            switch (BingImageColor)
            {
                case 0:
                    return "Color";
                // case 1:
                default:
                    return "Monochrome";
            }
        }

        private string GetBingImageStyle()
        {
            switch (BingImageStyle)
            {
                case 0:
                    return "Photo";
                //case 1:
                default:
                    return "Graphics";
            }
        }

        private string GetBingImageFace()
        {
            switch (BingImageFace)
            {
                case 0:
                    return "";
                case 1:
                    return "Face";
                case 2:
                    return "Portrait";
                // case 3:
                default:
                    return "Other";
            }
        }

        public string GetSearchEngine()
        {
            switch (SearchEngine)
            {
                case 0:
                    return TextCollection.ImagesLabText.SearchEngineBing;
                default:
                    return TextCollection.ImagesLabText.SearchEngineGoogle;
            }
        }

        #endregion
    }
}
