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

        private bool shiftIncludePosition = true;

        public bool ShiftIncludePosition
        {
            get { return shiftIncludePosition; }
            set
            {
                shiftIncludePosition = value;
                OnPropertyChanged("ShiftIncludePosition");
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

        private bool savedIncludePosition = true;

        public bool SavedIncludePosition
        {
            get { return savedIncludePosition; }
            set
            {
                savedIncludePosition = value;
                OnPropertyChanged("SavedIncludePosition");
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
            }
        }
        # endregion

        private Horizontal horizontalPosition;
        private Vertical verticalPosition;

        public Horizontal HorizontalPosition
        {
            get { return horizontalPosition; }
            set
            {
                horizontalPosition = value;
                OnPropertyChanged("PositionAlignment");
            }
        }

        public Vertical VerticalPosition
        {
            get { return verticalPosition; }
            set
            {
                verticalPosition = value;
                OnPropertyChanged("PositionAlignment");
            }
        }

        public Alignment PositionAlignment
        {
            get {
                switch (horizontalPosition)
                {
                    case Horizontal.Left:
                        switch (verticalPosition)
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
                        switch (verticalPosition)
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
                        switch (verticalPosition)
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
            set
            {
                switch (value)
                {
                    case Alignment.BottomLeft:
                        verticalPosition = Vertical.Bottom;
                        horizontalPosition = Horizontal.Left;
                        break;
                    case Alignment.BottomCenter:
                        verticalPosition = Vertical.Bottom;
                        horizontalPosition = Horizontal.Center;
                        break;
                    case Alignment.BottomRight:
                        verticalPosition = Vertical.Bottom;
                        horizontalPosition = Horizontal.Right;
                        break;
                    case Alignment.MiddleLeft:
                        verticalPosition = Vertical.Middle;
                        horizontalPosition = Horizontal.Left;
                        break;
                    case Alignment.MiddleCenter:
                        verticalPosition = Vertical.Middle;
                        horizontalPosition = Horizontal.Center;
                        break;
                    case Alignment.MiddleRight:
                        verticalPosition = Vertical.Middle;
                        horizontalPosition = Horizontal.Right;
                        break;
                    case Alignment.TopLeft:
                        verticalPosition = Vertical.Top;
                        horizontalPosition = Horizontal.Left;
                        break;
                    case Alignment.TopCenter:
                        verticalPosition = Vertical.Top;
                        horizontalPosition = Horizontal.Center;
                        break;
                    case Alignment.TopRight:
                        verticalPosition = Vertical.Top;
                        horizontalPosition = Horizontal.Right;
                        break;
                }
                OnPropertyChanged("PositionAlignment");
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
