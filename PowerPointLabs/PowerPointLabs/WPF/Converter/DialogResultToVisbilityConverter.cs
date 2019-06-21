using System;
using System.Windows;
using System.Windows.Data;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Windows;

namespace PowerPointLabs.WPF.Converter
{
    class DialogResultToVisbilityConverter : BaseConverter, IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value == null)
            {
                return DependencyProperty.UnsetValue;
            }
            return GetVisbility((DialogResult)value);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            // not implemented
            return value;
        }

        public static Visibility GetVisbility(DialogResult result)
        {
            return (result == DialogResult.None) ? Visibility.Hidden : Visibility.Visible;
        }
    }
}
