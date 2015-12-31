using PowerPointLabs.ImagesLab.Model;
using PowerPointLabs.ImagesLab.Service.Preview;

namespace PowerPointLabs.ImagesLab.Service.Interface
{
    interface IStylesDesigner
    {
        PreviewInfo PreviewApplyStyle(ImageItem source, StyleOptions option);
        void ApplyStyle(ImageItem source, StyleOptions option = null);
        void SetStyleOptions(StyleOptions opt);
        void CleanUp();
    }
}
