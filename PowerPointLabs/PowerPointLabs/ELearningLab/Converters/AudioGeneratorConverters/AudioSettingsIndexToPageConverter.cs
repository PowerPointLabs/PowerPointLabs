using System;
using System.Diagnostics;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Markup;
using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.ELearningLab.Views;

namespace PowerPointLabs.ELearningLab.Converters
{
    public class AudioSettingsIndexToPageConverter: MarkupExtension, IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // Find the appropriate page
            switch ((AudioSettingsPage)value)
            {
                case AudioSettingsPage.MainSettingsPage:
                    return AudioMainSettingsPage.GetInstance();
                case AudioSettingsPage.AzureLoginPage:
                    AzureVoiceLoginPage loginInstance = AzureVoiceLoginPage.GetInstance();
                    loginInstance.key.Text = "";
                    loginInstance.endpoint.SelectedIndex = -1;
                    return loginInstance;
                case AudioSettingsPage.AudioPreviewPage:
                    return AudioPreviewPage.GetInstance();
                default:
                    Debugger.Break();
                    return null;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return this;
        }
    }
}
