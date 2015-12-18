using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace PowerPointLabs.ImagesLab.Util
{
    class BoolToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var flag = value as bool? ?? false;
            return flag ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var vis = value as Visibility? ?? Visibility.Collapsed;
            return vis == Visibility.Visible;
        }
    }
}
