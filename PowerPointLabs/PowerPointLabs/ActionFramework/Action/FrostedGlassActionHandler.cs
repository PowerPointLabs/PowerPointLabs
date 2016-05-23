using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("frostedGlassEffect")]
    class FrostedGlassActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            var selection = this.GetCurrentSelection();
            var slide = this.GetCurrentSlide();
            var presentation = this.GetCurrentPresentation();
            FrostedGlass.FrostedGlassEffect(slide, presentation.SlideWidth, presentation.SlideHeight, selection);
        }
    }
}
