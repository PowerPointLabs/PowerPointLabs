using System;
using System.Drawing;
using System.Windows.Data;

using PowerPointLabs.ColorsLab;

namespace PowerPointLabs.Converters.ColorPane
{
    [ValueConversion(typeof(HSLColor), typeof(string))]
    class SelectedColorShiftLuminosityByFactor : IValueConverter
    {
        public static SelectedColorShiftLuminosityByFactor Instance = new SelectedColorShiftLuminosityByFactor();

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value == null)
            {
                Color color = (HSLColor)value;
                return ColorHelper.ColorToHexString(color);
            }
            HSLColor selectedColor = (HSLColor)value;
            float shiftFactor = float.Parse((string)parameter);
            Color convertedColor = new HSLColor(selectedColor.Hue, selectedColor.Saturation, shiftFactor * 240);
            return ColorHelper.ColorToHexString(convertedColor);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
