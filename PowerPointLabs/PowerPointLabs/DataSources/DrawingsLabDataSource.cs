using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using Microsoft.Office.Core;
using PowerPointLabs.Utils;

namespace PowerPointLabs.DataSources
{
    [SuppressMessage("Microsoft.StyleCop.CSharp.OrderingRules", "SA1202:ElementsMustBeOrderedByAccess", Justification = "To refactor to partials")]
    public class DrawingsLabDataSource : INotifyPropertyChanged
    {
        internal DrawingLabData Data { get; private set; }

        public DrawingsLabDataSource()
        {
            Data = new DrawingLabData();
            Data.AddPropertyChangedHandler(CallPropertyChanged);
        }

        internal void AssignData(DrawingLabData data)
        {
            this.Data = data;
            Data.AddPropertyChangedHandler(CallPropertyChanged);
            CallPropertyChanged("");
        }

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

        public bool HotkeysEnabled
        {
            get { return Data.HotkeysEnabled; }
            set
            {
                Data.HotkeysEnabled = value;
                OnPropertyChanged("HotkeysEnabled");
            }
        }

        # region Properties - Record / Apply Displacement
        public float ShiftValueX
        {
            get { return Data.ShiftValueX; }
            set
            {
                Data.ShiftValueX = value;
                OnPropertyChanged("ShiftValueX");
            }
        }

        public float ShiftValueY
        {
            get { return Data.ShiftValueY; }
            set
            {
                Data.ShiftValueY = value;
                OnPropertyChanged("ShiftValueY");
            }
        }

        public float ShiftValueRotation
        {
            get { return Data.ShiftValueRotation; }
            set
            {
                Data.ShiftValueRotation = value;
                OnPropertyChanged("ShiftValueRotation");
            }
        }

        public bool ShiftIncludePositionX
        {
            get { return Data.ShiftIncludePositionX; }
            set
            {
                Data.ShiftIncludePositionX = value;
                OnPropertyChanged("ShiftIncludePositionX");
            }
        }

        public bool ShiftIncludePositionY
        {
            get { return Data.ShiftIncludePositionY; }
            set
            {
                Data.ShiftIncludePositionY = value;
                OnPropertyChanged("ShiftIncludePositionY");
            }
        }

        public bool ShiftIncludeRotation
        {
            get { return Data.ShiftIncludeRotation; }
            set
            {
                Data.ShiftIncludeRotation = value;
                OnPropertyChanged("ShiftIncludeRotation");
            }
        }
        # endregion

        # region Properties - Record / Apply Position

        public float SavedValueX
        {
            get { return Data.SavedValueX; }
            set
            {
                Data.SavedValueX = value;
                OnPropertyChanged("SavedValueX");
            }
        }

        public float SavedValueY
        {
            get { return Data.SavedValueY; }
            set
            {
                Data.SavedValueY = value;
                OnPropertyChanged("SavedValueY");
            }
        }

        public float SavedValueRotation
        {
            get { return Data.SavedValueRotation; }
            set
            {
                Data.SavedValueRotation = value;
                OnPropertyChanged("SavedValueRotation");
            }
        }


        public bool SavedIncludePositionX
        {
            get { return Data.SavedIncludePositionX; }
            set
            {
                Data.SavedIncludePositionX = value;
                OnPropertyChanged("SavedIncludePositionX");
            }
        }

        public bool SavedIncludePositionY
        {
            get { return Data.SavedIncludePositionY; }
            set
            {
                Data.SavedIncludePositionY = value;
                OnPropertyChanged("SavedIncludePositionY");
            }
        }

        public bool SavedIncludeRotation
        {
            get { return Data.SavedIncludeRotation; }
            set
            {
                Data.SavedIncludeRotation = value;
                OnPropertyChanged("SavedIncludeRotation");
            }
        }

        # endregion

        # region Properties - Record / Apply Format

        public bool FormatSyncTextStyle
        {
            get { return Data.FormatSyncTextStyle; }
            set
            {
                Data.FormatSyncTextStyle = value;
                OnPropertyChanged("FormatSyncTextStyle");
            }
        }

        public string FormatText
        {
            get { return Data.FormatText; }
            set
            {
                Data.FormatText = value;
                OnPropertyChanged("FormatText");
            }
        }

        public bool FormatIncludeText
        {
            get { return Data.FormatIncludeText; }
            set
            {
                Data.FormatIncludeText = value;
                OnPropertyChanged("FormatIncludeText");
            }
        }

        public int FormatTextColor
        {
            get { return Data.FormatTextColor; }
            set
            {
                Data.FormatTextColor = value;
                OnPropertyChanged("FormatTextColor");
            }
        }

        public bool FormatIncludeTextColor
        {
            get { return Data.FormatIncludeTextColor; }
            set
            {
                Data.FormatIncludeTextColor = value;
                OnPropertyChanged("FormatIncludeTextColor");
            }
        }

        public float FormatTextFontSize
        {
            get { return Data.FormatTextFontSize; }
            set
            {
                Data.FormatTextFontSize = value;
                OnPropertyChanged("FormatTextFontSize");
            }
        }

        public bool FormatIncludeTextFontSize
        {
            get { return Data.FormatIncludeTextFontSize; }
            set
            {
                Data.FormatIncludeTextFontSize = value;
                OnPropertyChanged("FormatIncludeTextFontSize");
            }
        }

        public string FormatTextFont
        {
            get { return Data.FormatTextFont; }
            set
            {
                Data.FormatTextFont = value;
                OnPropertyChanged("FormatTextFont");
            }
        }

        public bool FormatIncludeTextFont
        {
            get { return Data.FormatIncludeTextFont; }
            set
            {
                Data.FormatIncludeTextFont = value;
                OnPropertyChanged("FormatIncludeTextFont");
            }
        }

        public bool FormatTextWrap
        {
            get { return Data.FormatTextWrap; }
            set
            {
                Data.FormatTextWrap = value;
                OnPropertyChanged("FormatTextWrap");
            }
        }

        public bool FormatIncludeTextWrap
        {
            get { return Data.FormatIncludeTextWrap; }
            set
            {
                Data.FormatIncludeTextWrap = value;
                OnPropertyChanged("FormatIncludeTextWrap");
            }
        }

        public MsoAutoSize FormatTextAutoSize
        {
            get { return Data.FormatTextAutoSize; }
            set
            {
                Data.FormatTextAutoSize = value;
                OnPropertyChanged("FormatTextAutoSize");
            }
        }

        public bool FormatIncludeTextAutoSize
        {
            get { return Data.FormatIncludeTextAutoSize; }
            set
            {
                Data.FormatIncludeTextAutoSize = value;
                OnPropertyChanged("FormatIncludeTextAutoSize");
            }
        }

        public bool FormatSyncLineStyle
        {
            get { return Data.FormatSyncLineStyle; }
            set
            {
                Data.FormatSyncLineStyle = value;
                OnPropertyChanged("FormatSyncLineStyle");
            }
        }

        public bool FormatHasLine
        {
            get { return Data.FormatHasLine; }
            set
            {
                Data.FormatHasLine = value;
                OnPropertyChanged("FormatHasLine");
            }
        }

        public bool FormatIncludeHasLine
        {
            get { return Data.FormatIncludeHasLine; }
            set
            {
                Data.FormatIncludeHasLine = value;
                OnPropertyChanged("FormatIncludeHasLine");
            }
        }

        public int FormatLineColor
        {
            get { return Data.FormatLineColor; }
            set
            {
                Data.FormatLineColor = value;
                OnPropertyChanged("FormatLineColor");
            }
        }

        public bool FormatIncludeLineColor
        {
            get { return Data.FormatIncludeLineColor; }
            set
            {
                Data.FormatIncludeLineColor = value;
                OnPropertyChanged("FormatIncludeLineColor");
            }
        }

        public float FormatLineWeight
        {
            get { return Data.FormatLineWeight; }
            set
            {
                Data.FormatLineWeight = value;
                OnPropertyChanged("FormatLineWeight");
            }
        }

        public bool FormatIncludeLineWeight
        {
            get { return Data.FormatIncludeLineWeight; }
            set
            {
                Data.FormatIncludeLineWeight = value;
                OnPropertyChanged("FormatIncludeLineWeight");
            }
        }

        public MsoLineDashStyle FormatLineDashStyle
        {
            get { return Data.FormatLineDashStyle; }
            set
            {
                Data.FormatLineDashStyle = value;
                OnPropertyChanged("FormatLineDashStyle");
            }
        }

        public bool FormatIncludeLineDashStyle
        {
            get { return Data.FormatIncludeLineDashStyle; }
            set
            {
                Data.FormatIncludeLineDashStyle = value;
                OnPropertyChanged("FormatIncludeLineDashStyle");
            }
        }

        public bool FormatSyncFillStyle
        {
            get { return Data.FormatSyncFillStyle; }
            set
            {
                Data.FormatSyncFillStyle = value;
                OnPropertyChanged("FormatSyncFillStyle");
            }
        }

        public bool FormatHasFill
        {
            get { return Data.FormatHasFill; }
            set
            {
                Data.FormatHasFill = value;
                OnPropertyChanged("FormatHasFill");
            }
        }

        public bool FormatIncludeHasFill
        {
            get { return Data.FormatIncludeHasFill; }
            set
            {
                Data.FormatIncludeHasFill = value;
                OnPropertyChanged("FormatIncludeHasFill");
            }
        }

        public int FormatFillColor
        {
            get { return Data.FormatFillColor; }
            set
            {
                Data.FormatFillColor = value;
                OnPropertyChanged("FormatFillColor");
            }
        }

        public bool FormatIncludeFillColor
        {
            get { return Data.FormatIncludeFillColor; }
            set
            {
                Data.FormatIncludeFillColor = value;
                OnPropertyChanged("FormatIncludeFillColor");
            }
        }

        public bool FormatSyncSize
        {
            get { return Data.FormatSyncSize; }
            set
            {
                Data.FormatSyncSize = value;
                OnPropertyChanged("FormatSyncSize");
            }
        }

        public float FormatWidth
        {
            get { return Data.FormatWidth; }
            set
            {
                Data.FormatWidth = value;
                OnPropertyChanged("FormatWidth");
            }
        }

        public bool FormatIncludeWidth
        {
            get { return Data.FormatIncludeWidth; }
            set
            {
                Data.FormatIncludeWidth = value;
                OnPropertyChanged("FormatIncludeWidth");
            }
        }

        public float FormatHeight
        {
            get { return Data.FormatHeight; }
            set
            {
                Data.FormatHeight = value;
                OnPropertyChanged("FormatHeight");
            }
        }

        public bool FormatIncludeHeight
        {
            get { return Data.FormatIncludeHeight; }
            set
            {
                Data.FormatIncludeHeight = value;
                OnPropertyChanged("FormatIncludeHeight");
            }
        }
        # endregion

        # region Properties - Anchor

        public Horizontal AnchorHorizontal
        {
            get { return Data.AnchorHorizontal; }
            set
            {
                Data.AnchorHorizontal = value;
                OnPropertyChanged("Anchor");
            }
        }

        public Vertical AnchorVertical
        {
            get { return Data.AnchorVertical; }
            set
            {
                Data.AnchorVertical = value;
                OnPropertyChanged("Anchor");
            }
        }

        public Alignment Anchor
        {
            get
            {
                return HorizontalVerticalToAlignment(Data.AnchorHorizontal, Data.AnchorVertical);
            }
            set
            {
                AlignmentToHorizontalVertical(value, out Data.AnchorHorizontal, out Data.AnchorVertical);
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
                    if (s.ToLower().StartsWith("mso"))
                    {
                        s = s.Substring(3);
                    }

                    if (s.ToLower().StartsWith("line"))
                    {
                        s = s.Substring(4);
                    }

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
            Data.PropertyChanged(propertyName);
        }

        private void CallPropertyChanged(string propertyName)
        {
            PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
        # endregion
    }
}
