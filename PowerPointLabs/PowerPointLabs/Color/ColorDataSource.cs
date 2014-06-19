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

        private Color themeColorOneValue;

        public Color themeColorOne
        {
            get
            {
                return themeColorOneValue;
            }
            set
            {
                if (value != this.themeColorOneValue)
                {
                    this.themeColorOneValue = value;
                    OnPropertyChanged("themeColorOne");
                }
            }
        }

        private Color themeColorTwoValue;

        public Color themeColorTwo
        {
            get
            {
                return themeColorTwoValue;
            }
            set
            {
                if (value != this.themeColorTwoValue)
                {
                    this.themeColorTwoValue = value;
                    OnPropertyChanged("themeColorTwo");
                }
            }
        }

        private Color themeColorThreeValue;

        public Color themeColorThree
        {
            get
            {
                return themeColorThreeValue;
            }
            set
            {
                if (value != this.themeColorThreeValue)
                {
                    this.themeColorThreeValue = value;
                    OnPropertyChanged("themeColorThree");
                }
            }
        }

        private Color themeColorFourValue;

        public Color themeColorFour
        {
            get
            {
                return themeColorFourValue;
            }
            set
            {
                if (value != this.themeColorFourValue)
                {
                    this.themeColorFourValue = value;
                    OnPropertyChanged("themeColorFour");
                }
            }
        }

        private Color themeColorFiveValue;

        public Color themeColorFive
        {
            get
            {
                return themeColorFiveValue;
            }
            set
            {
                if (value != this.themeColorFiveValue)
                {
                    this.themeColorFiveValue = value;
                    OnPropertyChanged("themeColorFive");
                }
            }
        }

        private Color themeColorSixValue;

        public Color themeColorSix
        {
            get
            {
                return themeColorSixValue;
            }
            set
            {
                if (value != this.themeColorSixValue)
                {
                    this.themeColorSixValue = value;
                    OnPropertyChanged("themeColorSix");
                }
            }
        }

        private Color themeColorSevenValue;

        public Color themeColorSeven
        {
            get
            {
                return themeColorSevenValue;
            }
            set
            {
                if (value != this.themeColorSevenValue)
                {
                    this.themeColorSevenValue = value;
                    OnPropertyChanged("themeColorSeven");
                }
            }
        }

        private Color themeColorEightValue;

        public Color themeColorEight
        {
            get
            {
                return themeColorEightValue;
            }
            set
            {
                if (value != this.themeColorEightValue)
                {
                    this.themeColorEightValue = value;
                    OnPropertyChanged("themeColorEight");
                }
            }
        }

        private Color themeColorNineValue;

        public Color themeColorNine
        {
            get
            {
                return themeColorNineValue;
            }
            set
            {
                if (value != this.themeColorNineValue)
                {
                    this.themeColorNineValue = value;
                    OnPropertyChanged("themeColorNine");
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
