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

            PowerPoint.Selection selection = this.GetCurrentSelection();

            if (currentSlide != null)
            {
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    foreach (PowerPoint.Shape selectedShape in this.GetAddIn().Application.ActiveWindow.Selection.ShapeRange)
                    {
                        CreateTooltip.GenerateCalloutWithReferenceTriggerShape(currentSlide, selectedShape);
                    }
                }
                else if (selection.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    PowerPoint.Shape triggerShape = CreateTooltip.GenerateTriggerShape(currentSlide);
                    PowerPoint.Shape callout = CreateTooltip.GenerateCalloutWithReferenceTriggerShape(currentSlide, triggerShape);
                }
            }
        }
    }
}
