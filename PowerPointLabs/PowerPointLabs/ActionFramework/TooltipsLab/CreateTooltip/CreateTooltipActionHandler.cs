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

                if (currentSlide != null)
                {
                    CreateTooltip.GenerateTriggerShape(currentSlide);
                }
            }
            catch (Exception)
            {
                if (currentSlide != null)
                {
                    CreateTooltip.GenerateCallout(currentSlide);
                    CreateTooltip.GenerateTriggerShape(currentSlide);
                }
            }
        }
    }
}
