using System;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.TooltipsLab;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ActionFramework.TooltipsLab
{
    [ExportActionRibbonId(TooltipsLabText.CreateTooltipTag)]
    class CreateTooltipActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            PowerPointSlide currentSlide = this.GetCurrentSlide();

            if (currentSlide == null)
            {
                return;
            }
            
            PowerPoint.Shape triggerShape = CreateTooltip.GenerateTriggerShape(currentSlide);
            PowerPoint.Shape callout = CreateTooltip.GenerateCalloutWithReferenceTriggerShape(currentSlide, triggerShape);
            PowerPoint.Shape calloutGroup = AddTextbox.AddTextboxToCallout(currentSlide, callout);
            AssignTooltip.AddTriggerAnimation(currentSlide, triggerShape, calloutGroup);

            if (!this.GetApplication().CommandBars.GetPressedMso(TooltipsLabConstants.AnimationPaneName))
            {
                this.GetApplication().CommandBars.ExecuteMso(TooltipsLabConstants.AnimationPaneName);
            }
        }
    }
}
