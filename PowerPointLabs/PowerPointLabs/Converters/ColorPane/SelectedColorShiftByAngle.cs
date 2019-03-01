using System;
using System.Drawing;
using System.Windows.Data;

using PowerPointLabs.ColorsLab;

namespace PowerPointLabs.Converters.ColorPane
{
    [ValueConversion(typeof(HSLColor), typeof(string))]
    class SelectedColorShiftByAngle : IValueConverter
    {
        public static SelectedColorShiftByAngle Instance = new SelectedColorShiftByAngle();

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value == null)
            {
                Color color = (HSLColor)value;
                return ColorHelper.ColorToHexString(color);
            }
            HSLColor selectedColor = (HSLColor)value;
            Color convertedColor = ColorHelper.GetColorShiftedByAngle(selectedColor, float.Parse((string)parameter));
            return ColorHelper.ColorToHexString(convertedColor);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
