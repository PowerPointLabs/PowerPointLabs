using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.Preview;

namespace PowerPointLabs.PictureSlidesLab.Service.Interface
{
    public interface IStylesDesigner
    {
        PreviewInfo PreviewApplyStyle(ImageItem source, Slide contentSlide, 
            float slideWidth, float slideHeight, StyleOption option);
        void ApplyStyle(ImageItem source, Slide contentSlide,
            float slideWidth, float slideHeight, StyleOption option = null);
        void SetStyleOptions(StyleOption opt);
        void SetSettings(Settings settings);
        void CleanUp();
    }
}
