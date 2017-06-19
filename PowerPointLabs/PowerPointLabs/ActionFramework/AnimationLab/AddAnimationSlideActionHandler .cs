using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.AnimationLab
{
    [ExportActionRibbonId("AddAnimationSlide")]
    class AddAnimationSlideActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            AutoAnimate.AddAutoAnimation();
        }
    }
}
