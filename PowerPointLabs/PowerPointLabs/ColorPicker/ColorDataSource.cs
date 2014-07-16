using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;

namespace PowerPointLabs
{
    class ColorDataSource : INotifyPropertyChanged
    {
        private bool isFillColorSelectedValue;

        public bool isFillColorSelected
        {
            get
            {
                return isFillColorSelectedValue;
            }
            set
            {
                if (value != this.isFillColorSelectedValue)
                {
                    this.isFillColorSelectedValue = value;
                    OnPropertyChanged("isFillColorSelected");
                }
            }
        }

        private bool isFontColorSelectedValue;

        public bool isFontColorSelected
        {
            get
            {
                return isFontColorSelectedValue;
            }
            set
            {
                if (value != this.isFontColorSelectedValue)
                {
                    this.isFontColorSelectedValue = value;
                    OnPropertyChanged("isFontColorSelected");
                }
            }
        }

        private bool isLineColorSelectedValue;

        public bool isLineColorSelected
        {
            get
            {
                return isLineColorSelectedValue;
            }
            set
            {
                if (value != this.isLineColorSelectedValue)
                {
                    this.isLineColorSelectedValue = value;
                    OnPropertyChanged("isLineColorSelected");
                }
            }
        }

        private HSLColor selectedColorValue;

        public HSLColor selectedColor
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

        private HSLColor themeColorOneValue;

        public HSLColor themeColorOne
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

        private HSLColor themeColorTwoValue;

        public HSLColor themeColorTwo
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

        private HSLColor themeColorThreeValue;

        public HSLColor themeColorThree
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

        private HSLColor themeColorFourValue;

        public HSLColor themeColorFour
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

        private HSLColor themeColorFiveValue;

        public HSLColor themeColorFive
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

        private HSLColor themeColorSixValue;

        public HSLColor themeColorSix
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

        private HSLColor themeColorSevenValue;

        public HSLColor themeColorSeven
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

        private HSLColor themeColorEightValue;

        public HSLColor themeColorEight
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

        private HSLColor themeColorNineValue;

        public HSLColor themeColorNine
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

        private HSLColor themeColorTenValue;

        public HSLColor themeColorTen
        {
            get
            {
                return themeColorTenValue;
            }
            set
            {
                if (value != this.themeColorTenValue)
                {
                    this.themeColorTenValue = value;
                    OnPropertyChanged("themeColorTen");
                }
            }
        }
        
        public bool SaveThemeColorsInFile(String filePath)
        {
            try
            {
                List<HSLColor> themeColors = new List<HSLColor>();
                themeColors.Add(this.themeColorOne);
                themeColors.Add(this.themeColorTwo);
                themeColors.Add(this.themeColorThree);
                themeColors.Add(this.themeColorFour);
                themeColors.Add(this.themeColorFive);
                themeColors.Add(this.themeColorSix);
                themeColors.Add(this.themeColorSeven);
                themeColors.Add(this.themeColorEight);
                themeColors.Add(this.themeColorNine);
                themeColors.Add(this.themeColorTen);

                Stream fileStream = File.Create(filePath);
                BinaryFormatter serializer = new BinaryFormatter();
                serializer.Serialize(fileStream, themeColors);
                fileStream.Close();
            }
            catch (Exception e)
            {
                return false;
            }
            return true;
        }

        public bool LoadThemeColorsFromFile(string filePath)
        {
            try
            {
                Stream openFileStream = File.OpenRead(filePath);
                BinaryFormatter deserializer = new BinaryFormatter();
                List<HSLColor> themeColors = (List<HSLColor>)deserializer.Deserialize(openFileStream);
                openFileStream.Close();

                this.themeColorOne = themeColors[0];
                this.themeColorTwo = themeColors[1];
                this.themeColorThree = themeColors[2];
                this.themeColorFour = themeColors[3];
                this.themeColorFive = themeColors[4];
                this.themeColorSix = themeColors[5];
                this.themeColorSeven = themeColors[6];
                this.themeColorEight = themeColors[7];
                this.themeColorNine = themeColors[8];
                this.themeColorTen = themeColors[9];
            }
            catch (Exception e)
            {
                return false;
            }
            return true;
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
