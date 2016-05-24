using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("EffectsLabFrostedGlassButton")]
    class EffectsLabFrostedGlassActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            Common.Log.Logger.Log("Entering Effects Lab Frosted Glass");

            this.StartNewUndoEntry();

            var selection = this.GetCurrentSelection();
            var slide = this.GetCurrentSlide();
            var presentation = this.GetCurrentPresentation();
            EffectsLabFrostedGlass.FrostedGlassEffect(slide, presentation.SlideWidth, presentation.SlideHeight, selection);
        }
    }
}
