using System.Globalization;
using System.Threading;
using System.Windows.Data;
using System.Windows.Forms;

namespace PowerPointLabs
{
    public class CustomBinding : System.Windows.Forms.Binding
    {
        private readonly IValueConverter _converter;
        private readonly object _converterParameter;
        private readonly CultureInfo _converterCulture;

        public CustomBinding(string propertyName, object dataSource, string dataMember, IValueConverter valueConverter, CultureInfo culture, object converterParameter = null)
            : base(propertyName, dataSource, dataMember)
        {
            if (valueConverter != null)
            {
                this._converter = valueConverter;
            }

            this.DataSourceUpdateMode = DataSourceUpdateMode.OnPropertyChanged;
            this.FormattingEnabled = false;

            this._converterCulture = culture;
            this._converterParameter = converterParameter;

        }

        public CustomBinding(string propertyName, object dataSource, string dataMember, IValueConverter valueConverter, object converterParameter = null)
            : base(propertyName, dataSource, dataMember)
        {
            if (valueConverter != null)
            {
                this._converter = valueConverter;
            }

            this._converterCulture = Thread.CurrentThread.CurrentUICulture;
            this._converterParameter = converterParameter;
        }



        protected override void OnFormat(ConvertEventArgs cevent)
        {
            if (this._converter != null)
            {
                object converterdValue = this._converter.Convert(cevent.Value, cevent.DesiredType, _converterParameter,
                                                              _converterCulture);
                cevent.Value = converterdValue;
            }
            else
            {
                base.OnFormat(cevent);
            }
        }

        protected override void OnParse(ConvertEventArgs cevent)
        {
            if (this._converter != null)
            {
                object converterdValue = this._converter.ConvertBack(cevent.Value, cevent.DesiredType, _converterParameter,
                                                                  _converterCulture);

                cevent.Value = converterdValue;
            }
            else
            {
                base.OnParse(cevent);
            }
        }
    }
}