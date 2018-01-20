﻿using System;
using System.Globalization;
using System.Windows.Data;

namespace PowerPointLabs.PictureSlidesLab.Util
{
    public class StyleVariationWidthConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            double width = (double) value;
            return width + 10;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            double actualWidth = (double) value;
            return actualWidth - 10;
        }
    }
}
