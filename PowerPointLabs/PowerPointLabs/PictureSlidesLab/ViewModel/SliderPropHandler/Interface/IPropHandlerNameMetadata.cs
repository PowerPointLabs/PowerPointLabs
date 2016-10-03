using System.ComponentModel;

namespace PowerPointLabs.PictureSlidesLab.ViewModel.SliderPropHandler.Interface
{
    public interface IPropHandlerNameMetadata
    {
        [DefaultValue("")]
        string PropHandlerName { get; }
    }
}
