using System.ComponentModel.Composition;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ViewModel.SliderPropHandler.Interface;

namespace PowerPointLabs.PictureSlidesLab.ViewModel.SliderPropHandler
{
    [Export(typeof(ISliderPropHandler))]
    [ExportMetadata("PropHandlerName", "Brightness")]
    class BrightnessSliderPropHandler : ISliderPropHandler
    {
        public Factory.SliderPropHandlerFactory.SliderProperties GetSliderProperties(StyleOption option)
        {
            var sliderProperties = new Factory.SliderPropHandlerFactory.SliderProperties();
            var colorValue = option.OverlayColor;
            var optValue = option.OverlayTransparency;
            sliderProperties.Value = (colorValue == "#FFFFFF") ? 200 - optValue : optValue;
            sliderProperties.Maximum = 200;
            sliderProperties.TickFrequency = 1;

            return sliderProperties;
        }

        public void BindStyleOption(StyleOption option, int value)
        {
            option.OptionName = "Customized";
            option.IsUseOverlayStyle = true;
            
            if (value > 100)
            {
                option.OverlayColor = "#FFFFFF";
                option.OverlayTransparency = 200 - value;
            }
            else
            {
                option.OverlayColor = "#000000";
                option.OverlayTransparency = value;
            }
        }

        public void BindStyleVariant(StyleVariant variant, int value)
        {
            variant.Set("OptionName", "Customized");
            variant.Set("IsUseOverlayStyle", true);
            
            if (value > 100)
            {
                variant.Set("OverlayColor", "#FFFFFF");
                variant.Set("OverlayTransparency", 200 - value);
            }
            else
            {
                variant.Set("OverlayColor", "#000000");
                variant.Set("OverlayTransparency", value);
            }
        }
    }
}

