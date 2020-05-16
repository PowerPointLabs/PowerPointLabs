using System;
using System.Windows;
using System.Windows.Data;

namespace PowerPointLabs.ELearningLab.Converters
{
    public class ShapeNameWidthConverter : IValueConverter
    {
        private const int leftIconWidth = 45;
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value is double)
            {
                return (double)value - leftIconWidth;
            }
            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value is double)
            {
                return (double)value + leftIconWidth;
            }
            return null;
        }
    }
}
