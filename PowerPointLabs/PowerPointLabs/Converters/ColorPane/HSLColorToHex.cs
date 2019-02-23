using System;
using System.Drawing;
using System.Windows.Data;

using PowerPointLabs.ColorsLab;

namespace PowerPointLabs.Converters.ColorPane
{
    [ValueConversion(typeof(HSLColor), typeof(string))]
    class HSLColorToHex : IValueConverter
    {
        public static HSLColorToHex Instance = new HSLColorToHex();

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            Color color = (HSLColor)value;
            return ColorHelper.ColorToHexString(color);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
