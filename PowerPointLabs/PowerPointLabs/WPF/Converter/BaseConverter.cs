using System;
using System.Windows.Markup;

namespace PowerPointLabs.WPF.Converter
{
    public abstract class BaseConverter : MarkupExtension
    {
        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return this;
        }
    }
}
