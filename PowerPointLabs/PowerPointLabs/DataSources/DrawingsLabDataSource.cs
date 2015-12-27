using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using PowerPointLabs.Utils;

namespace PowerPointLabs.DataSources
{
    public class DrawingsLabDataSource : INotifyPropertyChanged
    {
        public enum Alignment
        {
            TopLeft,
            TopCenter,
            TopRight,
            MiddleLeft,
            MiddleCenter,
            MiddleRight,
            BottomLeft,
            BottomCenter,
            BottomRight,
        }

        public enum Vertical
        {
            Top,
            Middle,
            Bottom,
        }

        public enum Horizontal
        {
            Left,
            Center,
            Right,
        }

        private bool _hotkeysEnabled = true;

        public bool HotkeysEnabled
        {
            get { return _hotkeysEnabled; }
            set
            {
                _hotkeysEnabled = value;
                OnPropertyChanged("HotkeysEnabled");
            }
        }

        # region Properties - Record / Apply Displacement
        private float _shiftValueX;
        private float _shiftValueY;
        private float _shiftValueRotation;
        private bool _shiftIncludePositionX = true;
        private bool _shiftIncludePositionY = true;
        private bool _shiftIncludeRotation = true;

        public float ShiftValueX
        {
            get { return _shiftValueX; }
            set
            {
                _shiftValueX = value;
                OnPropertyChanged("ShiftValueX");
            }
        }

        public float ShiftValueY
        {
            get { return _shiftValueY; }
            set
            {
                _shiftValueY = value;
                OnPropertyChanged("ShiftValueY");
            }
        }

        public float ShiftValueRotation
        {
            get { return _shiftValueRotation; }
            set
            {
                _shiftValueRotation = value;
                OnPropertyChanged("ShiftValueRotation");
            }
        }

        public bool ShiftIncludePositionX
        {
            get { return _shiftIncludePositionX; }
            set
            {
                _shiftIncludePositionX = value;
                OnPropertyChanged("ShiftIncludePositionX");
            }
        }

        public bool ShiftIncludePositionY
        {
            get { return _shiftIncludePositionY; }
            set
            {
                _shiftIncludePositionY = value;
                OnPropertyChanged("ShiftIncludePositionY");
            }
        }

        public bool ShiftIncludeRotation
        {
            get { return _shiftIncludeRotation; }
            set
            {
                _shiftIncludeRotation = value;
                OnPropertyChanged("ShiftIncludeRotation");
            }
        }
        # endregion

        # region Properties - Record / Apply Position
        private float _savedValueX;
        private float _savedValueY;
        private float _savedValueRotation;
        private bool _savedIncludePositionX = true;
        private bool _savedIncludePositionY = true;
        private bool _savedIncludeRotation = true;

        public float SavedValueX
        {
            get { return _savedValueX; }
            set
            {
                _savedValueX = value;
                OnPropertyChanged("SavedValueX");
            }
        }

        public float SavedValueY
        {
            get { return _savedValueY; }
            set
            {
                _savedValueY = value;
                OnPropertyChanged("SavedValueY");
            }
        }

        public float SavedValueRotation
        {
            get { return _savedValueRotation; }
            set
            {
                _savedValueRotation = value;
                OnPropertyChanged("SavedValueRotation");
            }
        }


        public bool SavedIncludePositionX
        {
            get { return _savedIncludePositionX; }
            set
            {
                _savedIncludePositionX = value;
                OnPropertyChanged("SavedIncludePositionX");
            }
        }

        public bool SavedIncludePositionY
        {
            get { return _savedIncludePositionY; }
            set
            {
                _savedIncludePositionY = value;
                OnPropertyChanged("SavedIncludePositionY");
            }
        }

        public bool SavedIncludeRotation
        {
            get { return _savedIncludeRotation; }
            set
            {
                _savedIncludeRotation = value;
                OnPropertyChanged("SavedIncludeRotation");
            }
        }

        # endregion

        # region Properties - Record / Apply Format
        private bool _formatSyncTextStyle = true;

        private string _formatText = "Text";
        private bool _formatIncludeText = false;
        private int _formatTextColor = 0x000000;
        private bool _formatIncludeTextColor = true;
        private float _formatTextFontSize = 5;
        private bool _formatIncludeTextFontSize = true;
        private string _formatTextFont = "Arial";
        private bool _formatIncludeTextFont = true;
        private bool _formatTextWrap = false;
        private bool _formatIncludeTextWrap = true;
        private MsoAutoSize _formatTextAutoSize = MsoAutoSize.msoAutoSizeNone;
        private bool _formatIncludeTextAutoSize = true;


        private bool _formatSyncLineStyle = true;

        private bool _formatHasLine = true;
        private bool _formatIncludeHasLine = true;
        private int _formatLineColor = 0x000000;
        private bool _formatIncludeLineColor = true;
        private float _formatLineWeight = 5;
        private bool _formatIncludeLineWeight = true;
        private MsoLineDashStyle _formatLineDashStyle = MsoLineDashStyle.msoLineSolid;
        private bool _formatIncludeLineDashStyle = true;


        private bool _formatSyncFillStyle = true;

        private bool _formatHasFill = true;
        private bool _formatIncludeHasFill = true;
        private int _formatFillColor = 0xC07000;
        private bool _formatIncludeFillColor = true;


        private bool _formatSyncSize = false;

        private float _formatWidth = 0;
        private bool _formatIncludeWidth = true;
        private float _formatHeight = 0;
        private bool _formatIncludeHeight = true;

        public bool FormatSyncTextStyle
        {
            get { return _formatSyncTextStyle; }
            set
            {
                _formatSyncTextStyle = value;
                OnPropertyChanged("FormatSyncTextStyle");
            }
        }

        public string FormatText
        {
            get { return _formatText; }
            set
            {
                _formatText = value;
                OnPropertyChanged("FormatText");
            }
        }

        public bool FormatIncludeText
        {
            get { return _formatIncludeText; }
            set
            {
                _formatIncludeText = value;
                OnPropertyChanged("FormatIncludeText");
            }
        }

        public int FormatTextColor
        {
            get { return _formatTextColor; }
            set
            {
                _formatTextColor = value;
                OnPropertyChanged("FormatTextColor");
            }
        }

        public bool FormatIncludeTextColor
        {
            get { return _formatIncludeTextColor; }
            set
            {
                _formatIncludeTextColor = value;
                OnPropertyChanged("FormatIncludeTextColor");
            }
        }

        public float FormatTextFontSize
        {
            get { return _formatTextFontSize; }
            set
            {
                _formatTextFontSize = value;
                OnPropertyChanged("FormatTextFontSize");
            }
        }

        public bool FormatIncludeTextFontSize
        {
            get { return _formatIncludeTextFontSize; }
            set
            {
                _formatIncludeTextFontSize = value;
                OnPropertyChanged("FormatIncludeTextFontSize");
            }
        }

        public string FormatTextFont
        {
            get { return _formatTextFont; }
            set
            {
                _formatTextFont = value;
                OnPropertyChanged("FormatTextFont");
            }
        }

        public bool FormatIncludeTextFont
        {
            get { return _formatIncludeTextFont; }
            set
            {
                _formatIncludeTextFont = value;
                OnPropertyChanged("FormatIncludeTextFont");
            }
        }

        public bool FormatTextWrap
        {
            get { return _formatTextWrap; }
            set
            {
                _formatTextWrap = value;
                OnPropertyChanged("FormatTextWrap");
            }
        }

        public bool FormatIncludeTextWrap
        {
            get { return _formatIncludeTextWrap; }
            set
            {
                _formatIncludeTextWrap = value;
                OnPropertyChanged("FormatIncludeTextWrap");
            }
        }

        public MsoAutoSize FormatTextAutoSize
        {
            get { return _formatTextAutoSize; }
            set
            {
                _formatTextAutoSize = value;
                OnPropertyChanged("FormatTextAutoSize");
            }
        }

        public bool FormatIncludeTextAutoSize
        {
            get { return _formatIncludeTextAutoSize; }
            set
            {
                _formatIncludeTextAutoSize = value;
                OnPropertyChanged("FormatIncludeTextAutoSize");
            }
        }

        public bool FormatSyncLineStyle
        {
            get { return _formatSyncLineStyle; }
            set
            {
                _formatSyncLineStyle = value;
                OnPropertyChanged("FormatSyncLineStyle");
            }
        }

        public bool FormatHasLine
        {
            get { return _formatHasLine; }
            set
            {
                _formatHasLine = value;
                OnPropertyChanged("FormatHasLine");
            }
        }

        public bool FormatIncludeHasLine
        {
            get { return _formatIncludeHasLine; }
            set
            {
                _formatIncludeHasLine = value;
                OnPropertyChanged("FormatIncludeHasLine");
            }
        }

        public int FormatLineColor
        {
            get { return _formatLineColor; }
            set
            {
                _formatLineColor = value;
                OnPropertyChanged("FormatLineColor");
            }
        }

        public bool FormatIncludeLineColor
        {
            get { return _formatIncludeLineColor; }
            set
            {
                _formatIncludeLineColor = value;
                OnPropertyChanged("FormatIncludeLineColor");
            }
        }

        public float FormatLineWeight
        {
            get { return _formatLineWeight; }
            set
            {
                _formatLineWeight = value;
                OnPropertyChanged("FormatLineWeight");
            }
        }

        public bool FormatIncludeLineWeight
        {
            get { return _formatIncludeLineWeight; }
            set
            {
                _formatIncludeLineWeight = value;
                OnPropertyChanged("FormatIncludeLineWeight");
            }
        }

        public MsoLineDashStyle FormatLineDashStyle
        {
            get { return _formatLineDashStyle; }
            set
            {
                _formatLineDashStyle = value;
                OnPropertyChanged("FormatLineDashStyle");
            }
        }

        public bool FormatIncludeLineDashStyle
        {
            get { return _formatIncludeLineDashStyle; }
            set
            {
                _formatIncludeLineDashStyle = value;
                OnPropertyChanged("FormatIncludeLineDashStyle");
            }
        }

        public bool FormatSyncFillStyle
        {
            get { return _formatSyncFillStyle; }
            set
            {
                _formatSyncFillStyle = value;
                OnPropertyChanged("FormatSyncFillStyle");
            }
        }

        public bool FormatHasFill
        {
            get { return _formatHasFill; }
            set
            {
                _formatHasFill = value;
                OnPropertyChanged("FormatHasFill");
            }
        }

        public bool FormatIncludeHasFill
        {
            get { return _formatIncludeHasFill; }
            set
            {
                _formatIncludeHasFill = value;
                OnPropertyChanged("FormatIncludeHasFill");
            }
        }

        public int FormatFillColor
        {
            get { return _formatFillColor; }
            set
            {
                _formatFillColor = value;
                OnPropertyChanged("FormatFillColor");
            }
        }

        public bool FormatIncludeFillColor
        {
            get { return _formatIncludeFillColor; }
            set
            {
                _formatIncludeFillColor = value;
                OnPropertyChanged("FormatIncludeFillColor");
            }
        }

        public bool FormatSyncSize
        {
            get { return _formatSyncSize; }
            set
            {
                _formatSyncSize = value;
                OnPropertyChanged("FormatSyncSize");
            }
        }

        public float FormatWidth
        {
            get { return _formatWidth; }
            set
            {
                _formatWidth = value;
                OnPropertyChanged("FormatWidth");
            }
        }

        public bool FormatIncludeWidth
        {
            get { return _formatIncludeWidth; }
            set
            {
                _formatIncludeWidth = value;
                OnPropertyChanged("FormatIncludeWidth");
            }
        }

        public float FormatHeight
        {
            get { return _formatHeight; }
            set
            {
                _formatHeight = value;
                OnPropertyChanged("FormatHeight");
            }
        }

        public bool FormatIncludeHeight
        {
            get { return _formatIncludeHeight; }
            set
            {
                _formatIncludeHeight = value;
                OnPropertyChanged("FormatIncludeHeight");
            }
        }
        # endregion

        # region Properties - Anchor
        private Horizontal _anchorHorizontal;
        private Vertical _anchorVertical;

        public Horizontal AnchorHorizontal
        {
            get { return _anchorHorizontal; }
            set
            {
                _anchorHorizontal = value;
                OnPropertyChanged("Anchor");
            }
        }

        public Vertical AnchorVertical
        {
            get { return _anchorVertical; }
            set
            {
                _anchorVertical = value;
                OnPropertyChanged("Anchor");
            }
        }

        public Alignment Anchor
        {
            get
            {
                return HorizontalVerticalToAlignment(_anchorHorizontal, _anchorVertical);
            }
            set
            {
                AlignmentToHorizontalVertical(value, out _anchorHorizontal, out _anchorVertical);
                OnPropertyChanged("Anchor");
            }
        }

        private static Alignment HorizontalVerticalToAlignment(Horizontal horizontal, Vertical vertical)
        {
            switch (horizontal)
            {
                case Horizontal.Left:
                    switch (vertical)
                    {
                        case Vertical.Top:
                            return Alignment.TopLeft;
                        case Vertical.Middle:
                            return Alignment.MiddleLeft;
                        case Vertical.Bottom:
                            return Alignment.BottomLeft;
                    }
                    break;
                case Horizontal.Center:
                    switch (vertical)
                    {
                        case Vertical.Top:
                            return Alignment.TopCenter;
                        case Vertical.Middle:
                            return Alignment.MiddleCenter;
                        case Vertical.Bottom:
                            return Alignment.BottomCenter;
                    }
                    break;
                case Horizontal.Right:
                    switch (vertical)
                    {
                        case Vertical.Top:
                            return Alignment.TopRight;
                        case Vertical.Middle:
                            return Alignment.MiddleRight;
                        case Vertical.Bottom:
                            return Alignment.BottomRight;
                    }
                    break;
            }
            throw new IndexOutOfRangeException();
        }

        private static void AlignmentToHorizontalVertical(Alignment alignment, out Horizontal horizontal, out Vertical vertical)
        {
            switch (alignment)
            {
                case Alignment.BottomLeft:
                    vertical = Vertical.Bottom;
                    horizontal = Horizontal.Left;
                    return;
                case Alignment.BottomCenter:
                    vertical = Vertical.Bottom;
                    horizontal = Horizontal.Center;
                    return;
                case Alignment.BottomRight:
                    vertical = Vertical.Bottom;
                    horizontal = Horizontal.Right;
                    return;
                case Alignment.MiddleLeft:
                    vertical = Vertical.Middle;
                    horizontal = Horizontal.Left;
                    return;
                case Alignment.MiddleCenter:
                    vertical = Vertical.Middle;
                    horizontal = Horizontal.Center;
                    return;
                case Alignment.MiddleRight:
                    vertical = Vertical.Middle;
                    horizontal = Horizontal.Right;
                    return;
                case Alignment.TopLeft:
                    vertical = Vertical.Top;
                    horizontal = Horizontal.Left;
                    return;
                case Alignment.TopCenter:
                    vertical = Vertical.Top;
                    horizontal = Horizontal.Center;
                    return;
                case Alignment.TopRight:
                    vertical = Vertical.Top;
                    horizontal = Horizontal.Right;
                    return;
                default:
                    throw new IndexOutOfRangeException();
            }
        }
        # endregion

        public Dictionary<MsoLineDashStyle, string> GetLineDashStyles
        {
            get
            {
                Func<MsoLineDashStyle, string> extractName = (style) =>
                {
                    var s = style.ToString();
                    if (s.ToLower().StartsWith("mso")) s = s.Substring(3);
                    if (s.ToLower().StartsWith("line")) s = s.Substring(4);
                    return Common.SplitCamelCase(s);
                };

                var styles = (MsoLineDashStyle[])Enum.GetValues(typeof(MsoLineDashStyle));
                return styles.ToDictionary(s => s, extractName);
            }
        }

        public Dictionary<MsoAutoSize, string> GetAutoSizeValues
        {
            get
            {
                return new Dictionary<MsoAutoSize, string>
                {
                    {MsoAutoSize.msoAutoSizeNone, "None"},
                    {MsoAutoSize.msoAutoSizeShapeToFitText, "Shape to fit Text"},
                    {MsoAutoSize.msoAutoSizeTextToFitShape, "Text to fit Shape"},
                };
            }
        }

        # region Event Implementation
        public event PropertyChangedEventHandler PropertyChanged = delegate {};

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
        # endregion
    }
}
