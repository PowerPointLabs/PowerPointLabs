using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using PowerPointLabs.ColorsLab;

namespace PowerPointLabs.DataSources
{
    class ColorDataSource : INotifyPropertyChanged
    {
        public IList<HSLColor> RecentColors
        {
            get
            {
                return recentColors;
            }
        }

        public IList<HSLColor> FavoriteColors
        {
            get
            {
                return favoriteColors;
            }
        }

        private readonly string[] recentColorFieldNames =
        {
            "RecentColorOne",
            "RecentColorTwo",
            "RecentColorThree",
            "RecentColorFour",
            "RecentColorFive",
            "RecentColorSix",
            "RecentColorSeven",
            "RecentColorEight",
            "RecentColorNine",
            "RecentColorTen",
            "RecentColorEleven",
            "RecentColorTwelve"
        };

        private readonly string[] favoriteColorFieldNames =
        {
            "favoriteColorOne",
            "favoriteColorTwo",
            "favoriteColorThree",
            "favoriteColorFour",
            "favoriteColorFive",
            "favoriteColorSix",
            "favoriteColorSeven",
            "favoriteColorEight",
            "favoriteColorNine",
            "favoriteColorTen",
            "favoriteColorEleven",
            "favoriteColorTwelve"
        };

        private ObservableCollection<HSLColor> recentColors;
        private ObservableCollection<HSLColor> favoriteColors;

        #region Properties

        private HSLColor selectedColorValue;

        public HSLColor SelectedColor
        {       
            get
            {
                return selectedColorValue;
            }
            set
            {
                if (value != selectedColorValue)
                {
                    selectedColorValue = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("selectedColor"));
                }
            }
        }

        #endregion

        #region Constructors

        public ColorDataSource()
        {
            recentColors = new ObservableCollection<HSLColor>();
            favoriteColors = new ObservableCollection<HSLColor>();
            recentColors.CollectionChanged += RecentColors_CollectionChanged;
            favoriteColors.CollectionChanged += FavoriteColors_CollectionChanged;
            ClearRecentColors();
            ClearFavoriteColors();
        }

        #endregion

        #region API

        public void AddColorToRecentColors(HSLColor color)
        {
            int index = recentColors.IndexOf(color);
            if (index == -1)
            {
                index = recentColors.Count - 1;
            }

            for (int i = index; i > 0; i--)
            {
                recentColors[i] = recentColors[i - 1];
            }
            recentColors[0] = color;
        }

        public void AddColorToFavorites(HSLColor color)
        {
            for (int i = favoriteColors.Count - 1; i > 0; i--)
            {
                favoriteColors[i] = favoriteColors[i - 1];
            }
            favoriteColors[0] = color;
        }

        public void ClearRecentColors()
        {
            recentColors.Clear();
            for (int i = 0; i < recentColorFieldNames.Length; i++)
            {
                recentColors.Add(Color.White);
            }
        }

        public void ClearFavoriteColors()
        {
            favoriteColors.Clear();
            for (int i = 0; i < favoriteColorFieldNames.Length; i++)
            {
                favoriteColors.Add(Color.White);
            }
        }

        public ObservableCollection<HSLColor> GetListOfRecentColors()
        {
            return recentColors;
        }

        public ObservableCollection<HSLColor> GetListOfFavoriteColors()
        {
            return favoriteColors;
        }

        public void SetRecentColor(int index, HSLColor color)
        {
            if (index >= recentColors.Count)
            {
                return;
            }
            recentColors[index] = color;
        }

        public void SetFavoriteColor(int index, HSLColor color)
        {
            if (index >= favoriteColors.Count)
            {
                return;
            }
            favoriteColors[index] = color;
        }

        #endregion

        #region Save/Load Colors

        public bool SaveRecentColorsInFile(string filePath)
        {
            try
            {
                Stream fileStream = File.Create(filePath);
                BinaryFormatter serializer = new BinaryFormatter();
                HSLColor[] colors = new HSLColor[recentColors.Count];
                recentColors.CopyTo(colors, 0);
                List<HSLColor> colorList = new List<HSLColor>(colors);
                serializer.Serialize(fileStream, colorList);
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
                List<HSLColor> newRecentColors = (List<HSLColor>)deserializer.Deserialize(openFileStream);
                openFileStream.Close();

                recentColors.Clear();
                foreach (HSLColor recentColor in newRecentColors)
                {
                    recentColors.Add(recentColor);
                }
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        public bool SaveFavoriteColorsInFile(String filePath)
        {
            try
            {
                Stream fileStream = File.Create(filePath);
                BinaryFormatter serializer = new BinaryFormatter();
                HSLColor[] colors = new HSLColor[favoriteColors.Count];
                favoriteColors.CopyTo(colors, 0);
                List<HSLColor> colorList = new List<HSLColor>(colors);
                serializer.Serialize(fileStream, colorList);
                fileStream.Close();
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        public bool LoadFavoriteColorsFromFile(string filePath)
        {
            try
            {
                Stream openFileStream = File.OpenRead(filePath);
                BinaryFormatter deserializer = new BinaryFormatter();
                List<HSLColor> newFavoriteColors = (List<HSLColor>)deserializer.Deserialize(openFileStream);
                openFileStream.Close();

                favoriteColors.Clear();
                foreach (HSLColor favoriteColor in newFavoriteColors)
                {
                    favoriteColors.Add(favoriteColor);
                }
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        #endregion

        #region EventHandlers

        public event PropertyChangedEventHandler PropertyChanged;

        // This method is called by the Set accessor of each property. 
        // The CallerMemberName attribute that is applied to the optional propertyName 
        // parameter causes the property name of the caller to be substituted as an argument. 
        // Create the OnPropertyChanged method to raise the event
        protected void RecentColors_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            System.Collections.IList newValues = e.NewItems;
            switch (newValues?.Count ?? 0)
            {
                case 0:
                    break;
                case 1:
                    string name = recentColorFieldNames[e.NewStartingIndex];
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
                    break;
                default:
                    foreach (string fieldName in recentColorFieldNames)
                    {
                        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(fieldName));
                    }
                    break;
            }
        }

        protected void FavoriteColors_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            System.Collections.IList newValues = e.NewItems;
            switch (newValues?.Count ?? 0)
            {
                case 0:
                    break;
                case 1:
                    string name = favoriteColorFieldNames[e.NewStartingIndex];
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
                    break;
                default:
                    foreach (string fieldName in favoriteColorFieldNames)
                    {
                        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(fieldName));
                    }
                    break;
            }
        }

        #endregion

    }
}
