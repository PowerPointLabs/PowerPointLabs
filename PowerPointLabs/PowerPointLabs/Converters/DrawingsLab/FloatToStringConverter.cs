using System;
using System.Globalization;
using System.Windows.Data;

namespace PowerPointLabs.Converters.DrawingsLab
{
    [ValueConversion(typeof(float), typeof(string))]
    class FloatToStringConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // return an invalid value in case of the value ends with a point
            return value.ToString().EndsWith(".") ? "." : value;
        }
    }
}
