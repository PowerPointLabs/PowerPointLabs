using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace PowerPointLabs.DataSources
{
    public class DrawingsLabAlignmentDataSource : INotifyPropertyChanged
    {
        public enum Position
        {
            Max,
            Mid,
            Min,
            None,
        }

        public Action sourcePropertyChangeEvent = () => { };
        public Action targetPropertyChangeEvent = () => { };

        private const float Max = 100;
        private const float Mid = 50;
        private const float Min = 0;

        private float sourceAnchor = Mid; // between 0 and 100
        private float targetAnchor = Mid; // between 0 and 100

        public float SourceAnchor
        {
            get { return sourceAnchor; }
            set
            {
                sourceAnchor = value;
                SourceOnPropertyChanged();
            }
        }

        public Position SourcePosition
        {
            get
            {
                return AnchorToPositionEnum(sourceAnchor);
            }
            set
            {
                PositionEnumToAnchor(value, ref sourceAnchor);
                SourceOnPropertyChanged();
            }
        }

        public float TargetAnchor
        {
            get { return targetAnchor; }
            set
            {
                targetAnchor = value;
                TargetOnPropertyChanged();
            }
        }

        public Position TargetPosition
        {
            get
            {
                return AnchorToPositionEnum(targetAnchor);
            }
            set
            {
                PositionEnumToAnchor(value, ref targetAnchor);
                TargetOnPropertyChanged();
            }
        }

        #region Event Implementation
        public event PropertyChangedEventHandler PropertyChanged = delegate { };

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
        # endregion

        #region Helper Functions
        private void PositionEnumToAnchor(Position value, ref float anchor)
        {
            switch (value)
            {
                case Position.Min:
                    anchor = Min;
                    break;
                case Position.Mid:
                    anchor = Mid;
                    break;
                case Position.Max:
                    anchor = Max;
                    break;
            }
        }

        private static Position AnchorToPositionEnum(float anchor)
        {
            if (anchor == Min)
            {
                return Position.Min;
            }
            if (anchor == Mid)
            {
                return Position.Mid;
            }
            if (anchor == Max)
            {
                return Position.Max;
            }
            return Position.None;
        }

        private void SourceOnPropertyChanged()
        {
            OnPropertyChanged("SourceAnchor");
            OnPropertyChanged("SourcePosition");
            sourcePropertyChangeEvent();
        }

        private void TargetOnPropertyChanged()
        {
            OnPropertyChanged("TargetAnchor");
            OnPropertyChanged("TargetPosition");
            targetPropertyChangeEvent();
        }
        #endregion
    }
}
