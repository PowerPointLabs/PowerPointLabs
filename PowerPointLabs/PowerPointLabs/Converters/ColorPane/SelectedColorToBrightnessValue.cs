using System;
using System.Windows.Data;

namespace PowerPointLabs.Converters.ColorPane
{
    class SelectedColorToBrightnessValue : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var selectedColor = (HSLColor)value;
            return (int)(selectedColor.Luminosity);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
