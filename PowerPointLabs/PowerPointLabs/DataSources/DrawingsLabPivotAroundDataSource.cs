using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace PowerPointLabs.DataSources
{
    public class DrawingsLabPivotAroundDataSource : INotifyPropertyChanged
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

        private int _copies = 1;
        private double _startAngle = 0;
        private double _angleDifference = 45;
        private bool _isExtend = false;
        private bool _fixOriginalLocation = true;
        private bool _rotateShape = true;
        private Alignment _pivotAnchor = Alignment.MiddleCenter;
        private Alignment _sourceAnchor = Alignment.MiddleCenter;

        public double StartAngle
        {
            get { return _startAngle; }
            set
            {
                _startAngle = value;
                OnPropertyChanged("StartAngle");
            }
        }

        public double AngleDifference
        {
            get { return _angleDifference; }
            set
            {
                _angleDifference = value;
                OnPropertyChanged("AngleDifference");
            }
        }

        public int Copies
        {
            get { return _copies; }
            set
            {
                _copies = value;
                OnPropertyChanged("Copies");
            }
        }

        public bool IsExtend
        {
            get { return _isExtend; }
            set
            {
                _isExtend = value;
                OnPropertyChanged("IsExtend");
            }
        }

        public bool FixOriginalLocation
        {
            get { return _fixOriginalLocation; }
            set
            {
                _fixOriginalLocation = value;
                OnPropertyChanged("FixOriginalLocation");
                OnPropertyChanged("StartAngleEditEnabled");
            }
        }

        public bool StartAngleEditEnabled
        {
            get { return !_fixOriginalLocation; }
        }

        public bool RotateShape
        {
            get { return _rotateShape; }
            set
            {
                _rotateShape = value;
                OnPropertyChanged("RotateShape");
            }
        }

        public Alignment PivotAnchor
        {
            get { return _pivotAnchor; }
            set
            {
                _pivotAnchor = value;
                OnPropertyChanged("PivotAnchor");
            }
        }

        public Alignment SourceAnchor
        {
            get { return _sourceAnchor; }
            set
            {
                _sourceAnchor = value;
                OnPropertyChanged("SourceAnchor");
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
