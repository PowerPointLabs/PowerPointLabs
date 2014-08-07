using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace PowerPointLabs.DataSources
{
    class ShapesLabSettingsDataSource : INotifyPropertyChanged
    {
        # region Properties
        private string _defaultSavingPath;
        public string DefaultSavingPath
        {
            get { return _defaultSavingPath; }
            set
            {
                _defaultSavingPath = value;
                OnPropertyChanged("DefaultSavingPath");
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
