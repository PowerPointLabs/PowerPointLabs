using System;
using System.Globalization;
using System.Windows.Data;

namespace PowerPointLabs.ImageSearch.Util
{
    public class StyleVariationWidthConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var width = (double) value;
            return width + 10;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var actualWidth = (double) value;
            return actualWidth - 10;
        }
    }
}
