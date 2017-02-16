using System;
using System.Globalization;
using System.Windows.Data;

namespace PowerPointLabs.Converters.DrawingsLab
{
    [ValueConversion(typeof(int), typeof(string))]
    class IntToStringConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value == null || string.IsNullOrEmpty(value.ToString())
                ? 0
                : value;
        }
    }
}
