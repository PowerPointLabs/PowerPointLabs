using System;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Data;

namespace PowerPointLabs.ELearningLab.Converters
{
    public class ShapeNameWidthConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value is double)
            {
                return (double)value - 45;
            }
            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value is double)
            {
                return (double)value + 45;
            }
            return null;
        }
    }
}
