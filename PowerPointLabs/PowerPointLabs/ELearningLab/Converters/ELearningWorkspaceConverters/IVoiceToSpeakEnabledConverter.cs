using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Markup;

using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.ELearningLab.Service;

namespace PowerPointLabs.ELearningLab.Converters
{
    public class IVoiceToSpeakEnabledConverter : MarkupExtension, IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            IVoice voice = value as IVoice;
            if (voice is ComputerVoice)
            {
                return true;
            }
            else
            {
                return AzureRuntimeService.IsAzureAccountPresentAndValid;
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
