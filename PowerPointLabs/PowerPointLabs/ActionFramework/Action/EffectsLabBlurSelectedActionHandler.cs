using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("EffectsLabBlurSelectedButton")]
    class EffectsLabBlurSelectedActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            Common.Log.Logger.Log("Entering Effects Lab Blur Selected");

            this.StartNewUndoEntry();

            var selection = this.GetCurrentSelection();
            var slide = this.GetCurrentSlide();
            var presentation = this.GetCurrentPresentation();
            EffectsLab.EffectsLabBlurSelected.BlurSelected(slide, selection);
        }
    }
}
