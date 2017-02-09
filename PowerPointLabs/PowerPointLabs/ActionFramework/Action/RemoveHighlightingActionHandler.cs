using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("RemoveHighlightButton")]
    class RemoveHighlightingHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.GetApplication().StartNewUndoEntry();
            var currentSlide = this.GetCurrentSlide();
            currentSlide.DeleteIndicator();
            currentSlide.DeleteShapesWithPrefix("PPTLabsHighlightBackgroundShape");
            currentSlide.DeleteShapesWithPrefix("PPTLabsHighlightTextFragmentsShape");
            foreach (Shape sh in currentSlide.Shapes)
            {
                if (sh.Name.Contains("HighlightTextShape"))
                {
                    currentSlide.DeleteShapeAnimations(sh);
                }
            }
        }
    }
}
