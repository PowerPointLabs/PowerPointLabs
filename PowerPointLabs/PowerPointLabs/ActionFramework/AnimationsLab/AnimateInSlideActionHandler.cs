using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.AnimationsLab
{
    [ExportActionRibbonId("AnimateInSlide")]
    class AnimateInSlideActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            AnimateInSlide.isHighlightBullets = false;
            AnimateInSlide.AddAnimationInSlide();
        }
    }
}
