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

            try
            {
                PowerPoint.ShapeRange selectedShapes = this.GetAddIn().Application.ActiveWindow.Selection.ShapeRange;

                // Maybe this check should be changed to do it in such a way that if there's no current slide,
                // simply grey out the button??
                if (currentSlide != null)
                {
                    foreach (PowerPoint.Shape selectedShape in selectedShapes)
                    {
                        CreateTooltip.GenerateCalloutWithReferenceTriggerShape(currentSlide, selectedShape);

                    }

                    // TODO: After generating the callout and trigger shapes, trigger animation should go here.
                }
            }
            catch (Exception)
            {
                if (currentSlide != null)
                {
                    PowerPoint.Shape triggerShape = CreateTooltip.GenerateTriggerShape(currentSlide);
                    CreateTooltip.GenerateCalloutWithReferenceTriggerShape(currentSlide, triggerShape);

                }
            }
        }
    }
}
