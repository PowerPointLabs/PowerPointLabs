using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;

namespace PowerPointLabs
{
    class ColorDataSource : INotifyPropertyChanged
    {
        private Color selectedColorValue;

        public Color selectedColor
        {       
            get 
            { 
                return selectedColorValue; 
            }
            set
            {
                if (value != this.selectedColorValue)
                {
                    this.selectedColorValue = value;
                    OnPropertyChanged("selectedColor");
                }
            }
        }
        

        public ColorDataSource()
        {
        }

        public event PropertyChangedEventHandler PropertyChanged;

        // This method is called by the Set accessor of each property. 
        // The CallerMemberName attribute that is applied to the optional propertyName 
        // parameter causes the property name of the caller to be substituted as an argument. 
        // Create the OnPropertyChanged method to raise the event 
        protected void OnPropertyChanged(string name)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(name));
            }
        }
    }
}
