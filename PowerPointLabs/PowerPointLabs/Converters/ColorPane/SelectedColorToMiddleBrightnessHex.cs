using System;
using System.Drawing;
using System.Windows.Data;

namespace PowerPointLabs.Converters.ColorPane
{
    [ValueConversion(typeof(HSLColor), typeof(string))]
    class SelectedColorToMiddleBrightnessHex : IValueConverter
    {

        public static SelectedColorToMiddleBrightnessHex Instance = new SelectedColorToMiddleBrightnessHex();

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value == null)
            {
                Color color = (HSLColor)value;
                return "#" + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
            }
            HSLColor selectedColor = (HSLColor)value;
            HSLColor midBrightnessHslColor = new HSLColor();
            midBrightnessHslColor.Hue = selectedColor.Hue;
            midBrightnessHslColor.Saturation = selectedColor.Saturation;
            midBrightnessHslColor.Luminosity = 120;
            Color midBrightnessColor = midBrightnessHslColor;
            return "#" + midBrightnessColor.R.ToString("X2") + midBrightnessColor.G.ToString("X2") + midBrightnessColor.B.ToString("X2");
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
