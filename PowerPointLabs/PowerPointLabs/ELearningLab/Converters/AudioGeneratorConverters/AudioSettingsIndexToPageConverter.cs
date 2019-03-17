using System;
using System.Diagnostics;
using System.Globalization;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Markup;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.ELearningLab.Views;

namespace PowerPointLabs.ELearningLab.Converters
{
    public class AudioSettingsIndexToPageConverter : MarkupExtension, IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            bool shouldGoToMainPage;
            Page mainPage, subPage;
            try
            {
                shouldGoToMainPage = (bool)values[0];
                mainPage = (Page)values[1];
                subPage = (Page)values[2];

                return shouldGoToMainPage ? mainPage : subPage;
            }
            catch
            {
                Logger.Log("Error converting binded value to appropriate pages");
                return null;
            }
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return this;
        }
    }
}
