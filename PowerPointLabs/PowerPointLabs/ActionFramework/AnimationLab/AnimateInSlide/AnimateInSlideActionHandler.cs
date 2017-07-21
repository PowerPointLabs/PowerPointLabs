using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.AnimationLab;

namespace PowerPointLabs.ActionFramework.AnimationLab
{
    [ExportActionRibbonId(TextCollection1.AnimateInSlideTag)]
    class AnimateInSlideActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();
            
            AnimateInSlide.AddAnimationInSlide();
        }
    }
}
