using PowerPointLabs.PictureSlidesLab.Model;

namespace PowerPointLabs.PictureSlidesLab.ViewModel.SliderPropHandler.Interface
{
    interface ISliderPropHandler
    {
        Factory.SliderPropHandlerFactory.SliderProperties GetSliderProperties(StyleOption option);
        void BindStyleOption(StyleOption option, int value);
        void BindStyleVariant(StyleVariant variant, int value);
    }
}
