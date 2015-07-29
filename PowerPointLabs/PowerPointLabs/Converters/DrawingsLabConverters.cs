using System;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;
using PowerPointLabs.ColorPicker;
using Color = System.Windows.Media.Color;

namespace PowerPointLabs.Converters
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

    public class EnumMatchToBooleanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType,
                              object parameter, CultureInfo culture)
        {
            if (value == null || parameter == null)
                return false;

            string checkValue = value.ToString();
            string targetValue = parameter.ToString();
            return checkValue.Equals(targetValue,
                     StringComparison.InvariantCultureIgnoreCase);
        }

        public object ConvertBack(object value, Type targetType,
                                  object parameter, CultureInfo culture)
        {
            if (value == null || parameter == null)
                return null;

            bool useValue = (bool)value;
            string targetValue = parameter.ToString();
            if (useValue)
                return Enum.Parse(targetType, targetValue);

            return null;
        }
    }

    public class ColorToBackgroundConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            byte r, g, b;
            Utils.Graphics.UnpackRgbInt((int) value, out r, out g, out b);
            return new SolidColorBrush(Color.FromRgb(r,g,b));
            
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var brush = value as SolidColorBrush;
            var color = brush.Color;
            return Utils.Graphics.PackRgbInt(color.R, color.G, color.B);
        }
    }
}
