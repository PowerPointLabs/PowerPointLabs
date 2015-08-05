using System;
using System.IO;
using System.Xml.Serialization;
using ImageProcessor.Imaging.Filters;
using PowerPointLabs.ImageSearch.Handler.Effect;
using PowerPointLabs.WPF.Observable;

namespace PowerPointLabs.ImageSearch.Domain
{
    [Serializable]
    public class StyleOptions : Model
    {
        public StyleOptions()
        {
            Init();
        }

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

        private int _textBoxPosition;

        public int TextBoxPosition
        {
            get { return _textBoxPosition; }
            set
            {
                _textBoxPosition = value;
                OnPropertyChanged("TextBoxPosition");
            }
        }

        private int _textBoxAlignment;

        public int TextBoxAlignment
        {
            get { return _textBoxAlignment; }
            set
            {
                _textBoxAlignment = value;
                OnPropertyChanged("TextBoxAlignment");
            }
        }

        // for direct text/special effect/blur style
        // for background's
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

        // for background's overlay
        public int Transparency
        {
            get { return _transparency; }
            set
            {
                _transparency = value;
                OnPropertyChanged("Transparency");
            }
        }

        // for textbox style

        private string _textBoxOverlayColor;

        public string TextBoxOverlayColor
        {
            get { return _textBoxOverlayColor; }
            set
            {
                _textBoxOverlayColor = value;
                OnPropertyChanged("TextBoxOverlayColor");
            }
        }

        private int _textBoxTransparency;

        public int TextBoxTransparency
        {
            get { return _textBoxTransparency; }
            set
            {
                _textBoxTransparency = value;
                OnPropertyChanged("TextBoxTransparency");
            }
        }

        // for banner style
        private int _bannerShape;

        public int BannerShape
        {
            get { return _bannerShape; }
            set
            {
                _bannerShape = value;
                OnPropertyChanged("BannerShape");
            }
        }

        private int _bannerDirection;

        public int BannerDirection
        {
            get { return _bannerDirection; }
            set
            {
                _bannerDirection = value;
                OnPropertyChanged("BannerDirection");
            }
        }

        private string _bannerOverlayColor;

        public string BannerOverlayColor
        {
            get { return _bannerOverlayColor; }
            set
            {
                _bannerOverlayColor = value;
                OnPropertyChanged("BannerOverlayColor");
            }
        }

        private int _bannerTransparency;

        public int BannerTransparency
        {
            get { return _bannerTransparency; }
            set
            {
                _bannerTransparency = value;
                OnPropertyChanged("BannerTransparency");
            }
        }

        // for special effect style
        private int _specialEffect;

        public int SpecialEffect
        {
            get { return _specialEffect; }
            set
            {
                _specialEffect = value;
                OnPropertyChanged("SpecialEffect");
            }
        }

        // other
        private bool _isInsertReference;

        public bool IsInsertReference
        {
            get { return _isInsertReference; }
            set
            {
                _isInsertReference = value;
                OnPropertyChanged("IsInsertReference");
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
            TextBoxPosition = 4;
            TextBoxAlignment = 0;

            OverlayColor = "#000000";
            Transparency = 85;
            BannerOverlayColor = "#000000";
            BannerTransparency = 25;
            TextBoxOverlayColor = "#D74926"; // red-orange
            TextBoxTransparency = 25;

            BannerShape = 0;
            BannerDirection = 0;
            
            SpecialEffect = 0;

            IsInsertReference = false;
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
                    return "Trebuchet MS";
                case 5:
                    return "";
                default:
                    return "Segoe UI";
            }
        }

        public Position GetTextBoxPosition()
        {
            switch (TextBoxPosition)
            {
                case 0:
                    return Position.Original;
                case 1:
                    return Position.TopLeft;
                case 2:
                    return Position.Top;
                case 3:
                    return Position.TopRight;
                case 4:
                    return Position.Left;
                case 5:
                    return Position.Centre;
                case 6:
                    return Position.Right;
                case 7:
                    return Position.BottomLeft;
                case 8:
                    return Position.Bottom;
                // case 9:
                default:
                    return Position.BottomRight;
            }
        }

        public Alignment GetTextBoxAlignment()
        {
            switch (TextBoxAlignment)
            {
                case 0:
                    return Alignment.Auto;
                case 1:
                    return Alignment.Left;
                case 2:
                    return Alignment.Centre;
                // case 3:
                default:
                    return Alignment.Right;
            }
        }

        public BannerShape GetBannerShape()
        {
            switch (BannerShape)
            {
                case 0:
                    return Handler.Effect.BannerShape.Rectangle;
                // case 1:
                default:
                    return Handler.Effect.BannerShape.Circle;
            }
        }

        public BannerDirection GetBannerDirection()
        {
            switch (BannerDirection)
            {
                case 0:
                    return Handler.Effect.BannerDirection.Auto;
                case 1:
                    return Handler.Effect.BannerDirection.Horizontal;
                // case 2:
                default:
                    return Handler.Effect.BannerDirection.Vertical;
            }
        }

        public IMatrixFilter GetSpecialEffect()
        {
            switch (SpecialEffect)
            {
                case 0:
                    return MatrixFilters.GreyScale;
                case 1:
                    return MatrixFilters.BlackWhite;
                case 2:
                    return MatrixFilters.Comic;
                case 3:
                    return MatrixFilters.Gotham;
                case 4:
                    return MatrixFilters.HiSatch;
                case 5:
                    return MatrixFilters.Invert;
                case 6:
                    return MatrixFilters.Lomograph;
                case 7:
                    return MatrixFilters.LoSatch;
                case 8:
                    return MatrixFilters.Polaroid;
                // case 9:
                default:
                    return MatrixFilters.Sepia;
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
