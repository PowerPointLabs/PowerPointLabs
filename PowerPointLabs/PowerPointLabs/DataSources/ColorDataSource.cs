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
        private const int numRecentColors = 12;
        private const int numFavoriteColors = 12;

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

            for (int i = index - 1; i > 0; i--)
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
            for (int i = 0; i < numRecentColors; i++)
            {
                recentColors.Add(Color.White);
            }
        }

        public void ClearFavoriteColors()
        {
            favoriteColors.Clear();
            for (int i = 0; i < numFavoriteColors; i++)
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
                ObservableCollection<HSLColor> recentColors = GetListOfRecentColors();

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
                List<HSLColor> newRecentColors = (List<HSLColor>)deserializer.Deserialize(openFileStream);
                ObservableCollection<HSLColor> observableRecentColors = new ObservableCollection<HSLColor>(newRecentColors);
                openFileStream.Close();

                recentColors = observableRecentColors;
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
                serializer.Serialize(fileStream, favoriteColors);
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

                this.favoriteColors = new ObservableCollection<HSLColor>(newFavoriteColors);
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        #endregion

        #region EventHandlers

        public delegate void ColorChange(object sender, int colorIndex);

        public event ColorChange RecentColorChanged;
        public event ColorChange FavoriteColorChanged;

        public event PropertyChangedEventHandler PropertyChanged;

        // This method is called by the Set accessor of each property. 
        // The CallerMemberName attribute that is applied to the optional propertyName 
        // parameter causes the property name of the caller to be substituted as an argument. 
        // Create the OnPropertyChanged method to raise the event 
        protected void RecentColors_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            List<HSLColor> newValues = e.NewItems as List<HSLColor> ?? new List<HSLColor>();
            switch (newValues.Count)
            {
                case 0:
                    break;
                case 1:
                    RecentColorChanged?.Invoke(newValues[0], e.NewStartingIndex);
                    break;
                default:
                    for (int i = 0; i < recentColors.Count; i++)
                    {
                        RecentColorChanged?.Invoke(recentColors[i], i);
                    }
                    break;
            }
        }
        // This method is called by the Set accessor of each property. 
        // The CallerMemberName attribute that is applied to the optional propertyName 
        // parameter causes the property name of the caller to be substituted as an argument. 
        // Create the OnPropertyChanged method to raise the event 
        protected void FavoriteColors_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            List<HSLColor> newValues = e.NewItems as List<HSLColor> ?? new List<HSLColor>();
            switch (newValues.Count)
            {
                case 0:
                    break;
                case 1:
                    FavoriteColorChanged?.Invoke(newValues[0], e.NewStartingIndex);
                    break;
                default:
                    for (int i = 0; i < favoriteColors.Count; i++)
                    {
                        FavoriteColorChanged?.Invoke(favoriteColors[i], i);
                    }
                    break;
            }
        }

        #endregion

    }
}
