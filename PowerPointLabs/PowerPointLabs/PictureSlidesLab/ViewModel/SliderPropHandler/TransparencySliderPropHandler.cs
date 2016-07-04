using System.ComponentModel.Composition;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ViewModel.SliderPropHandler.Interface;

namespace PowerPointLabs.PictureSlidesLab.ViewModel.SliderPropHandler
{
    [Export(typeof(ISliderPropHandler))]
    [ExportMetadata("PropHandlerName", "Transparency")]
    class TransparencySliderPropHandler : ISliderPropHandler
    {
        public string PropName { get; set; }

        public Factory.SliderPropHandlerFactory.SliderProperties GetSliderProperties(StyleOption option)
        {
            var sliderProperties = new Factory.SliderPropHandlerFactory.SliderProperties();
            var type = option.GetType();
            var prop = type.GetProperty(PropName);
            sliderProperties.Value = (int)prop.GetValue(option, null);
            sliderProperties.Maximum = 100;
            sliderProperties.TickFrequency = 1;

            return sliderProperties;
        }

        public void BindStyleOption(StyleOption option, int value)
        {
            option.OptionName = "Customized";
            var type = option.GetType();
            var prop = type.GetProperty(PropName);
            prop.SetValue(option, value, null);
        }

        public void BindStyleVariant(StyleVariant variant, int value)
        {
            variant.Set("OptionName", "Customized");
            variant.Set(PropName, value);
        }
    }
}

