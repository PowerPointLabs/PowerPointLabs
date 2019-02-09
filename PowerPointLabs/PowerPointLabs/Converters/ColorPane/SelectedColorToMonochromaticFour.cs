using System;
using System.Drawing;
using System.Windows.Data;

namespace PowerPointLabs.Converters.ColorPane
{
    [ValueConversion(typeof(HSLColor), typeof(string))]
    class SelectedColorToMonochromaticFour : IValueConverter
    {
        public static SelectedColorToMonochromaticFour Instance = new SelectedColorToMonochromaticFour();

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value == null)
            {
                Color color = (HSLColor)value;
                return "#" + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
            }
            HSLColor selectedColor = (HSLColor)value;
            Color convertedColor = new HSLColor(selectedColor.Hue, selectedColor.Saturation, 0.50f * 240);
            return "#" + convertedColor.R.ToString("X2") + convertedColor.G.ToString("X2") + convertedColor.B.ToString("X2");
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
