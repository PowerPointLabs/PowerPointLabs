using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace PowerPointLabs.DataSources
{
    public class DrawingsLabMultiCloneGridDataSource : INotifyPropertyChanged
    {
        private static bool _isExtend = true;
        private static int _xCopies = 5;
        private static int _yCopies = 5;

        public int XCopies
        {
            get { return _xCopies; }
            set
            {
                _xCopies = value;
                OnPropertyChanged("XCopies");
            }
        }

        public int YCopies
        {
            get { return _yCopies; }
            set
            {
                _yCopies = value;
                OnPropertyChanged("YCopies");
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

        # region Event Implementation
        public event PropertyChangedEventHandler PropertyChanged = delegate {};

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
        # endregion
    }
}
