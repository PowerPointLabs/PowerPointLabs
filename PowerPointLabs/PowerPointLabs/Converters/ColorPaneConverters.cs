using System;
using System.Drawing;
using System.Windows.Data;
using PowerPointLabs.ColorPicker;

namespace PowerPointLabs.Converters
{
    class HSLColorToRGBColor : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            Color selectedColor = (HSLColor) value;
            return Color.FromArgb(255,
                    selectedColor.R,
                    selectedColor.G,
                    selectedColor.B);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    class selectedColorToAnalogousLower : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var selectedColor = (HSLColor)value;
            return ColorHelper.GetColorShiftedByAngle(selectedColor, -30.0f);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    class selectedColorToAnalogousHigher : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var selectedColor = (HSLColor)value;
            return ColorHelper.GetColorShiftedByAngle(selectedColor, 30.0f);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    class selectedColorToComplementaryColor : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var selectedColor = (HSLColor)value;
            return ColorHelper.GetColorShiftedByAngle(selectedColor, 180.0f);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    class selectedColorToSplitComplementaryLower : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var selectedColor = (HSLColor)value;
            return ColorHelper.GetColorShiftedByAngle(selectedColor, 150.0f);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    class selectedColorToSplitComplementaryHigher : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var selectedColor = (HSLColor)value;
            return ColorHelper.GetColorShiftedByAngle(selectedColor, 210.0f);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    class selectedColorToTriadicLower : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var selectedColor = (HSLColor)value;
            return ColorHelper.GetColorShiftedByAngle(selectedColor, -120.0f);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    class selectedColorToTriadicHigher : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var selectedColor = (HSLColor)value;
            return ColorHelper.GetColorShiftedByAngle(selectedColor, 120.0f);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    class selectedColorToTetradicOne : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var selectedColor = (HSLColor)value;
            return ColorHelper.GetColorShiftedByAngle(selectedColor, 90.0f);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    class selectedColorToTetradicTwo : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var selectedColor = (HSLColor)value;
            return ColorHelper.GetColorShiftedByAngle(selectedColor, 180.0f);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    class selectedColorToTetradicThree : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var selectedColor = (HSLColor)value;
            return ColorHelper.GetColorShiftedByAngle(selectedColor, 270.0f);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    class selectedColorToMonochromaticOne : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var selectedColor = (HSLColor) value;
            Color convertedColor = new HSLColor(selectedColor.Hue, selectedColor.Saturation, 0.80f*240);
            return Color.FromArgb(255,
                convertedColor.R,
                convertedColor.G,
                convertedColor.B);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    class selectedColorToMonochromaticTwo : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var selectedColor = (HSLColor) value;
            Color convertedColor = new HSLColor(selectedColor.Hue, selectedColor.Saturation, 0.70f * 240);
            return Color.FromArgb(255,
                convertedColor.R,
                convertedColor.G,
                convertedColor.B);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    class selectedColorToMonochromaticThree : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var selectedColor = (HSLColor)value;
            Color convertedColor = new HSLColor(selectedColor.Hue, selectedColor.Saturation, 0.60f * 240);
            return Color.FromArgb(255,
                convertedColor.R,
                convertedColor.G,
                convertedColor.B);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    class selectedColorToMonochromaticFour : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var selectedColor = (HSLColor)value;
            Color convertedColor = new HSLColor(selectedColor.Hue, selectedColor.Saturation, 0.50f * 240);
            return Color.FromArgb(255,
                convertedColor.R,
                convertedColor.G,
                convertedColor.B);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    class selectedColorToMonochromaticFive : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var selectedColor = (HSLColor)value;
            Color convertedColor = new HSLColor(selectedColor.Hue, selectedColor.Saturation, 0.40f * 240);
            return Color.FromArgb(255,
                convertedColor.R,
                convertedColor.G,
                convertedColor.B);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    class selectedColorToMonochromaticSix : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var selectedColor = (HSLColor)value;
            Color convertedColor = new HSLColor(selectedColor.Hue, selectedColor.Saturation, 0.30f * 240);
            return Color.FromArgb(255,
                convertedColor.R,
                convertedColor.G,
                convertedColor.B);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    class selectedColorToMonochromaticSeven : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var selectedColor = (HSLColor)value;
            Color convertedColor = new HSLColor(selectedColor.Hue, selectedColor.Saturation, 0.20f * 240);
            return Color.FromArgb(255,
                convertedColor.R,
                convertedColor.G,
                convertedColor.B);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    class selectedColorToBrightnessValue : IValueConverter
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
    class selectedColorToSaturationValue : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var selectedColor = (HSLColor)value;
            return (int)(selectedColor.Saturation);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    class IsActiveBoolToButtonBackColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            bool isActive = (bool)value;
            if (isActive)
            {
                return SystemColors.ActiveCaption;
            }
            return SystemColors.Control;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
