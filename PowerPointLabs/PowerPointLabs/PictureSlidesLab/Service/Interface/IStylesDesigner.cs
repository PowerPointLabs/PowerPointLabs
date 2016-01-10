using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.Preview;

namespace PowerPointLabs.PictureSlidesLab.Service.Interface
{
    interface IStylesDesigner
    {
        PreviewInfo PreviewApplyStyle(ImageItem source, Slide contentSlide, 
            float slideWidth, float slideHeight, StyleOptions option);
        void ApplyStyle(ImageItem source, Slide contentSlide, StyleOptions option = null);
        void SetStyleOptions(StyleOptions opt);
        void CleanUp();
    }
}
