using System.ComponentModel.Composition;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ViewModel.SliderPropHandler.Interface;

namespace PowerPointLabs.PictureSlidesLab.ViewModel.SliderPropHandler
{
    [Export(typeof(ISliderPropHandler))]
    [ExportMetadata("PropHandlerName", "Blurriness")]
    class BlurrinessSliderPropHandler : ISliderPropHandler
    {
        public Factory.SliderPropHandlerFactory.SliderProperties GetSliderProperties(StyleOption option)
        {
            var sliderProperties = new Factory.SliderPropHandlerFactory.SliderProperties();
            var optValue = option.BlurDegree;
            sliderProperties.Value = (optValue == 0) ? 0 : (optValue - 50) * 2;
            sliderProperties.Maximum = 100;
            sliderProperties.TickFrequency = 2;

            return sliderProperties;
        }

        public void BindStyleOption(StyleOption option, int value)
        {
            option.OptionName = "Customized";
            option.IsUseBlurStyle = true;
            option.BlurDegree = (value == 0) ? 0 : 50 + (value / 2);
        }

        public void BindStyleVariant(StyleVariant variant, int value)
        {
            variant.Set("OptionName", "Customized");
            variant.Set("IsUseBlurStyle", true);
            var variantValue = (value == 0) ? 0 : 50 + (value / 2);
            variant.Set("BlurDegree", variantValue);
        }
    }
}
