using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("CropToSlideButton")]
    class CropToSlideActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            CropLab.CropToSlide.Crop(this.GetCurrentSelection(), this.GetCurrentPresentation().SlideWidth, this.GetCurrentPresentation().SlideHeight);
        }
    }
}
