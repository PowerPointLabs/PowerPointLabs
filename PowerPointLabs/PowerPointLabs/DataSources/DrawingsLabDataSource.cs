using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

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

        # region Properties
        private float shiftValueX;

        public float ShiftValueX
        {
            get { return shiftValueX; }
            set
            {
                shiftValueX = value;
                OnPropertyChanged("ShiftValueX");
            }
        }

        private float shiftValueY;

        public float ShiftValueY
        {
            get { return shiftValueY; }
            set
            {
                shiftValueY = value;
                OnPropertyChanged("ShiftValueY");
            }
        }

        private float shiftValueRotation;

        public float ShiftValueRotation
        {
            get { return shiftValueRotation; }
            set
            {
                shiftValueRotation = value;
                OnPropertyChanged("ShiftValueRotation");
            }
        }

        private bool shiftIncludePositionX = true;

        public bool ShiftIncludePositionX
        {
            get { return shiftIncludePositionX; }
            set
            {
                shiftIncludePositionX = value;
                OnPropertyChanged("ShiftIncludePositionX");
                OnPropertyChanged("ShiftIncludePositionBoth");
            }
        }

        private bool shiftIncludePositionY = true;

        public bool ShiftIncludePositionY
        {
            get { return shiftIncludePositionY; }
            set
            {
                shiftIncludePositionY = value;
                OnPropertyChanged("ShiftIncludePositionY");
                OnPropertyChanged("ShiftIncludePositionBoth");
            }
        }

        public bool? ShiftIncludePositionBoth
        {
            get
            {
                if (ShiftIncludePositionX != ShiftIncludePositionY) return null;
                return ShiftIncludePositionX;
            }
            set
            {
                if (value == null) return;
                bool valueBool = (value == true);

                shiftIncludePositionX = valueBool;
                shiftIncludePositionY = valueBool;
                OnPropertyChanged("ShiftIncludePositionX");
                OnPropertyChanged("ShiftIncludePositionY");
                OnPropertyChanged("ShiftIncludePositionBoth");
            }
        }

        private bool shiftIncludeRotation = true;

        public bool ShiftIncludeRotation
        {
            get { return shiftIncludeRotation; }
            set
            {
                shiftIncludeRotation = value;
                OnPropertyChanged("ShiftIncludeRotation");
            }
        }

        private float savedValueX;

        public float SavedValueX
        {
            get { return savedValueX; }
            set
            {
                savedValueX = value;
                OnPropertyChanged("SavedValueX");
            }
        }

        private float savedValueY;

        public float SavedValueY
        {
            get { return savedValueY; }
            set
            {
                savedValueY = value;
                OnPropertyChanged("SavedValueY");
            }
        }

        private float savedValueRotation;

        public float SavedValueRotation
        {
            get { return savedValueRotation; }
            set
            {
                savedValueRotation = value;
                OnPropertyChanged("SavedValueRotation");
            }
        }

        private bool savedIncludePositionX = true;

        public bool SavedIncludePositionX
        {
            get { return savedIncludePositionX; }
            set
            {
                savedIncludePositionX = value;
                OnPropertyChanged("SavedIncludePositionX");
                OnPropertyChanged("SavedIncludePositionBoth");
            }
        }

        private bool savedIncludePositionY = true;

        public bool SavedIncludePositionY
        {
            get { return savedIncludePositionY; }
            set
            {
                savedIncludePositionY = value;
                OnPropertyChanged("SavedIncludePositionY");
                OnPropertyChanged("SavedIncludePositionBoth");
            }
        }

        public bool? SavedIncludePositionBoth
        {
            get
            {
                if (SavedIncludePositionX != SavedIncludePositionY) return null;
                return SavedIncludePositionX;
            }
            set
            {
                if (value == null) return;
                bool valueBool = (value == true);

                savedIncludePositionX = valueBool;
                savedIncludePositionY = valueBool;
                OnPropertyChanged("SavedIncludePositionX");
                OnPropertyChanged("SavedIncludePositionY");
                OnPropertyChanged("SavedIncludePositionBoth");
            }
        }

        private bool savedIncludeRotation = true;

        public bool SavedIncludeRotation
        {
            get { return savedIncludeRotation; }
            set
            {
                savedIncludeRotation = value;
                OnPropertyChanged("SavedIncludeRotation");
                OnPropertyChanged("SavedIncludePositionBoth");
            }
        }

        private int _formatFillColor = 0xC07000;

        public int FormatFillColor
        {
            get
            {
                return _formatFillColor;
            }
            set
            {
                if (value == _formatFillColor) return;
                
                _formatFillColor = value;
                OnPropertyChanged("FormatFillColor");
            }
        }

        private int _formatLineColor = 0x000000;

        public int FormatLineColor
        {
            get
            {
                return _formatLineColor;
            }
            set
            {
                if (value == _formatLineColor) return;

                _formatLineColor = value;
                OnPropertyChanged("FormatLineColor");
            }
        }

        private float _formatLineWeight = 5;

        public float FormatLineWeight
        {
            get { return _formatLineWeight; }
            set
            {
                _formatLineWeight = value;
                OnPropertyChanged("FormatLineWeight");
            }
        }

        private bool formatIncludeFillColor = true;

        public bool FormatIncludeFillColor
        {
            get { return formatIncludeFillColor; }
            set
            {
                formatIncludeFillColor = value;
                OnPropertyChanged("FormatIncludeFillColor");
                OnPropertyChanged("FormatIncludeAll");
            }
        }

        private bool formatIncludeLineColor = true;

        public bool FormatIncludeLineColor
        {
            get { return formatIncludeLineColor; }
            set
            {
                formatIncludeLineColor = value;
                OnPropertyChanged("FormatIncludeLineColor");
                OnPropertyChanged("FormatIncludeAll");
            }
        }

        private bool formatIncludeLineWeight = true;

        public bool FormatIncludeLineWeight
        {
            get { return formatIncludeLineWeight; }
            set
            {
                formatIncludeLineWeight = value;
                OnPropertyChanged("FormatIncludeLineWeight");
                OnPropertyChanged("FormatIncludeAll");
            }
        }

        public bool? FormatIncludeAll
        {
            get
            {
                // If not all equal, return null.
                if (!(FormatIncludeFillColor == FormatIncludeLineColor &&
                      FormatIncludeLineColor == FormatIncludeLineWeight)) return null;
                return FormatIncludeFillColor;
            }
            set
            {
                if (value == null) return;
                bool valueBool = (value == true);

                formatIncludeFillColor = valueBool;
                formatIncludeLineColor = valueBool;
                formatIncludeLineWeight = valueBool;
                OnPropertyChanged("FormatIncludeFillColor");
                OnPropertyChanged("FormatIncludeLineColor");
                OnPropertyChanged("FormatIncludeLineWeight");
                OnPropertyChanged("FormatIncludeAll");
            }
        }

        private bool hotkeysEnabled = false;

        public bool HotkeysEnabled
        {
            get { return hotkeysEnabled; }
            set
            {
                hotkeysEnabled = value;
                OnPropertyChanged("HotkeysEnabled");
            }
        }

        # endregion

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

        # region Event Implementation
        public event PropertyChangedEventHandler PropertyChanged = delegate {};

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
        # endregion
    }
}
