
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.TooltipsLab;
using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ActionFramework.TooltipsLab
{
    [ExportActionRibbonId(TooltipsLabText.CreateCalloutTag)]
    class CreateCalloutActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();
            PowerPointSlide currentSlide = this.GetCurrentSlide();
            PowerPoint.Selection selection = this.GetCurrentSelection();

            if (currentSlide == null || !ShapeUtil.IsSelectionShape(selection))
            {
                return;
            }

            foreach (PowerPoint.Shape selectedShape in selection.ShapeRange)
            {
                PowerPoint.Shape callout = CreateTooltip.GenerateCalloutWithReferenceTriggerShape(currentSlide, selectedShape);

                ConvertToTooltip.AddTriggerAnimation(currentSlide, selectedShape, callout);
           }
            
            if (!this.GetApplication().CommandBars.GetPressedMso(TooltipsLabConstants.AnimationPaneName))
            {
                this.GetApplication().CommandBars.ExecuteMso(TooltipsLabConstants.AnimationPaneName);
            }
        }
    }
}
