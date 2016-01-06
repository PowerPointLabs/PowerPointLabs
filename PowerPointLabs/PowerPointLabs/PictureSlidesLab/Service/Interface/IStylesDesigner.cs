using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.Preview;

namespace PowerPointLabs.PictureSlidesLab.Service.Interface
{
    interface IStylesDesigner
    {
        PreviewInfo PreviewApplyStyle(ImageItem source, StyleOptions option);
        void ApplyStyle(ImageItem source, StyleOptions option = null);
        void SetStyleOptions(StyleOptions opt);
        void CleanUp();
    }
}
