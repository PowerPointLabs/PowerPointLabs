using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace PowerPointLabs.DataSources
{
    public class DrawingsLabDataSource : INotifyPropertyChanged
    {
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

        # region Event Implementation
        public event PropertyChangedEventHandler PropertyChanged = delegate {};

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
        # endregion
    }
}
