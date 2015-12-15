using System;
using System.Globalization;
using System.Windows.Data;

namespace PowerPointLabs.ImageSearch.Util
{
    class BoolToBackgroundColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var flag = value as bool? ?? false;
            return flag ? "#D74926" : "#454545";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
