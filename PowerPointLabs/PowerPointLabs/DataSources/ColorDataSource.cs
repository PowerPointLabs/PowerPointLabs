using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;

using PowerPointLabs.ColorsLab;

namespace PowerPointLabs.DataSources
{
    class ColorDataSource : INotifyPropertyChanged
    {
        private bool isFillColorSelectedValue;

        public bool IsFillColorSelected
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

        public bool IsFontColorSelected
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

        public bool IsLineColorSelected
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

        public HSLColor SelectedColor
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

        private HSLColor recentColorOneValue;

        public HSLColor RecentColorOne
        {
            get
            {
                return recentColorOneValue;
            }
            set
            {
                if (value != this.recentColorOneValue)
                {
                    this.recentColorOneValue = value;
                    OnPropertyChanged("RecentColorOne");
                }
            }
        }

        private HSLColor recentColorTwoValue;

        public HSLColor RecentColorTwo
        {
            get
            {
                return recentColorTwoValue;
            }
            set
            {
                if (value != this.recentColorTwoValue)
                {
                    this.recentColorTwoValue = value;
                    OnPropertyChanged("RecentColorTwo");
                }
            }
        }

        private HSLColor recentColorThreeValue;

        public HSLColor RecentColorThree
        {
            get
            {
                return recentColorThreeValue;
            }
            set
            {
                if (value != this.recentColorThreeValue)
                {
                    this.recentColorThreeValue = value;
                    OnPropertyChanged("RecentColorThree");
                }
            }
        }

        private HSLColor recentColorFourValue;

        public HSLColor RecentColorFour
        {
            get
            {
                return recentColorFourValue;
            }
            set
            {
                if (value != this.recentColorFourValue)
                {
                    this.recentColorFourValue = value;
                    OnPropertyChanged("RecentColorFour");
                }
            }
        }

        private HSLColor recentColorFiveValue;

        public HSLColor RecentColorFive
        {
            get
            {
                return recentColorFiveValue;
            }
            set
            {
                if (value != this.recentColorFiveValue)
                {
                    this.recentColorFiveValue = value;
                    OnPropertyChanged("RecentColorFive");
                }
            }
        }

        private HSLColor recentColorSixValue;

        public HSLColor RecentColorSix
        {
            get
            {
                return recentColorSixValue;
            }
            set
            {
                if (value != this.recentColorSixValue)
                {
                    this.recentColorSixValue = value;
                    OnPropertyChanged("RecentColorSix");
                }
            }
        }

        private HSLColor recentColorSevenValue;

        public HSLColor RecentColorSeven
        {
            get
            {
                return recentColorSevenValue;
            }
            set
            {
                if (value != this.recentColorSevenValue)
                {
                    this.recentColorSevenValue = value;
                    OnPropertyChanged("RecentColorSeven");
                }
            }
        }

        private HSLColor recentColorEightValue;

        public HSLColor RecentColorEight
        {
            get
            {
                return recentColorEightValue;
            }
            set
            {
                if (value != this.recentColorEightValue)
                {
                    this.recentColorEightValue = value;
                    OnPropertyChanged("RecentColorEight");
                }
            }
        }

        private HSLColor recentColorNineValue;

        public HSLColor RecentColorNine
        {
            get
            {
                return recentColorNineValue;
            }
            set
            {
                if (value != this.recentColorNineValue)
                {
                    this.recentColorNineValue = value;
                    OnPropertyChanged("RecentColorNine");
                }
            }
        }

        private HSLColor recentColorTenValue;

        public HSLColor RecentColorTen
        {
            get
            {
                return recentColorTenValue;
            }
            set
            {
                if (value != this.recentColorTenValue)
                {
                    this.recentColorTenValue = value;
                    OnPropertyChanged("RecentColorTen");
                }
            }
        }

        private HSLColor recentColorElevenValue;

        public HSLColor RecentColorEleven
        {
            get
            {
                return recentColorElevenValue;
            }
            set
            {
                if (value != this.recentColorElevenValue)
                {
                    this.recentColorElevenValue = value;
                    OnPropertyChanged("RecentColorEleven");
                }
            }
        }

        private HSLColor recentColorTwelveValue;

        public HSLColor RecentColorTwelve
        {
            get
            {
                return recentColorTwelveValue;
            }
            set
            {
                if (value != this.recentColorTwelveValue)
                {
                    this.recentColorTwelveValue = value;
                    OnPropertyChanged("RecentColorTwelve");
                }
            }
        }

        private HSLColor themeColorOneValue;

        public HSLColor ThemeColorOne
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

        public HSLColor ThemeColorTwo
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

        public HSLColor ThemeColorThree
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

        public HSLColor ThemeColorFour
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

        public HSLColor ThemeColorFive
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

        public HSLColor ThemeColorSix
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

        public HSLColor ThemeColorSeven
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

        public HSLColor ThemeColorEight
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

        public HSLColor ThemeColorNine
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

        public HSLColor ThemeColorTen
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

        private HSLColor themeColorElevenValue;

        public HSLColor ThemeColorEleven
        {
            get
            {
                return themeColorElevenValue;
            }
            set
            {
                if (value != this.themeColorElevenValue)
                {
                    this.themeColorElevenValue = value;
                    OnPropertyChanged("themeColorEleven");
                }
            }
        }

        private HSLColor themeColorTwelveValue;

        public HSLColor ThemeColorTwelve
        {
            get
            {
                return themeColorTwelveValue;
            }
            set
            {
                if (value != this.themeColorTwelveValue)
                {
                    this.themeColorTwelveValue = value;
                    OnPropertyChanged("themeColorTwelve");
                }
            }
        }

        private HSLColor themeColorThirteenValue;

        public HSLColor ThemeColorThirteen
        {
            get
            {
                return themeColorThirteenValue;
            }
            set
            {
                if (value != this.themeColorThirteenValue)
                {
                    this.themeColorThirteenValue = value;
                    OnPropertyChanged("themeColorThirteen");
                }
            }
        }

        // TODO: To improve: make it smarter
        //       When adding color, check if it is already in recents. If it is, remove it and put it in front.
        //       Therefore we will not have repeated colors in recent colors -- doesn't make sense.
        public void AddColorToRecentColors(HSLColor color)
        {
            RecentColorTwelve = RecentColorEleven;
            RecentColorEleven = RecentColorTen;
            RecentColorTen = RecentColorNine;
            RecentColorNine = RecentColorEight;
            RecentColorEight = RecentColorSeven;
            RecentColorSeven = RecentColorSix;
            RecentColorSix = RecentColorFive;
            RecentColorFive = RecentColorFour;
            RecentColorFour = RecentColorThree;
            RecentColorThree = RecentColorTwo;
            RecentColorTwo = RecentColorOne;
            RecentColorOne = color;
        }

        public bool SaveRecentColorsInFile(string filePath)
        {
            try
            {
                List<HSLColor> recentColors = new List<HSLColor>();
                recentColors.Add(this.RecentColorOne);
                recentColors.Add(this.RecentColorTwo);
                recentColors.Add(this.RecentColorThree);
                recentColors.Add(this.RecentColorFour);
                recentColors.Add(this.RecentColorFive);
                recentColors.Add(this.RecentColorSix);
                recentColors.Add(this.RecentColorSeven);
                recentColors.Add(this.RecentColorEight);
                recentColors.Add(this.RecentColorNine);
                recentColors.Add(this.RecentColorTen);
                recentColors.Add(this.RecentColorEleven);
                recentColors.Add(this.RecentColorTwelve);

                Stream fileStream = File.Create(filePath);
                BinaryFormatter serializer = new BinaryFormatter();
                serializer.Serialize(fileStream, recentColors);
                fileStream.Close();
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        public bool LoadRecentColorsFromFile(string filePath)
        {
            try
            {
                Stream openFileStream = File.OpenRead(filePath);
                BinaryFormatter deserializer = new BinaryFormatter();
                List<HSLColor> recentColors = (List<HSLColor>)deserializer.Deserialize(openFileStream);
                openFileStream.Close();

                this.RecentColorOne = recentColors[0];
                this.RecentColorTwo = recentColors[1];
                this.RecentColorThree = recentColors[2];
                this.RecentColorFour = recentColors[3];
                this.RecentColorFive = recentColors[4];
                this.RecentColorSix = recentColors[5];
                this.RecentColorSeven = recentColors[6];
                this.RecentColorEight = recentColors[7];
                this.RecentColorNine = recentColors[8];
                this.RecentColorTen = recentColors[9];
                this.RecentColorEleven = recentColors[10];
                this.RecentColorTwelve = recentColors[11];
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        public bool SaveThemeColorsInFile(String filePath)
        {
            try
            {
                List<HSLColor> themeColors = new List<HSLColor>();
                themeColors.Add(this.ThemeColorOne);
                themeColors.Add(this.ThemeColorTwo);
                themeColors.Add(this.ThemeColorThree);
                themeColors.Add(this.ThemeColorFour);
                themeColors.Add(this.ThemeColorFive);
                themeColors.Add(this.ThemeColorSix);
                themeColors.Add(this.ThemeColorSeven);
                themeColors.Add(this.ThemeColorEight);
                themeColors.Add(this.ThemeColorNine);
                themeColors.Add(this.ThemeColorTen);
                themeColors.Add(this.ThemeColorEleven);
                themeColors.Add(this.ThemeColorTwelve);
                themeColors.Add(this.ThemeColorThirteen);

                Stream fileStream = File.Create(filePath);
                BinaryFormatter serializer = new BinaryFormatter();
                serializer.Serialize(fileStream, themeColors);
                fileStream.Close();
            }
            catch (Exception)
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

                this.ThemeColorOne = themeColors[0];
                this.ThemeColorTwo = themeColors[1];
                this.ThemeColorThree = themeColors[2];
                this.ThemeColorFour = themeColors[3];
                this.ThemeColorFive = themeColors[4];
                this.ThemeColorSix = themeColors[5];
                this.ThemeColorSeven = themeColors[6];
                this.ThemeColorEight = themeColors[7];
                this.ThemeColorNine = themeColors[8];
                this.ThemeColorTen = themeColors[9];
                this.ThemeColorEleven = themeColors[10];
                this.ThemeColorTwelve = themeColors[11];
                this.ThemeColorThirteen = themeColors[12];
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        public void AddColorToFavorites(HSLColor color)
        {
            ThemeColorThirteen = ThemeColorTwelve;
            ThemeColorTwelve = ThemeColorEleven;
            ThemeColorEleven = ThemeColorTen;
            ThemeColorTen = ThemeColorNine;
            ThemeColorNine = ThemeColorEight;
            ThemeColorEight = ThemeColorSeven;
            ThemeColorSeven = ThemeColorSix;
            ThemeColorSix = ThemeColorFive;
            ThemeColorFive = ThemeColorFour;
            ThemeColorFour = ThemeColorThree;
            ThemeColorThree = ThemeColorTwo;
            ThemeColorTwo = ThemeColorOne;
            ThemeColorOne = color;
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
