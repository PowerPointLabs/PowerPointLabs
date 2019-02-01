using System;
using System.Diagnostics;
using System.Globalization;

using PowerPointLabs.NarrationsLab.Data;
using PowerPointLabs.NarrationsLab.Views;

namespace PowerPointLabs.NarrationsLab.ValueConverters
{
    public class NarrationsLabSettingsPageValueConverter: BaseValueConverter<NarrationsLabSettingsPageValueConverter>
    {
        public override object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // Find the appropriate page
            switch ((NarrationsLabSettingsPage)value)
            {
                case NarrationsLabSettingsPage.MainSettingsPage:
                    return NarrationsLabMainSettingsPage.GetInstance();
                case NarrationsLabSettingsPage.LoginPage:
                    AzureVoiceLoginPage loginInstance = AzureVoiceLoginPage.GetInstance();
                    loginInstance.key.Text = "";
                    loginInstance.endpoint.SelectedIndex = -1;
                    return loginInstance;             
                default:
                    Debugger.Break();
                    return null;
            }
        }

        public override object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
