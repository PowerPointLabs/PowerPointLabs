using System;
using System.Windows;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.TooltipsLab;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ActionFramework.TooltipsLab
{
    [ExportActionRibbonId(TooltipsLabText.CreateTriggerTag)]
    class CreateTriggerActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            PowerPointSlide currentSlide = this.GetCurrentSlide();

            if (currentSlide == null)
            {
                return;
            }

            PowerPoint.Selection selection = this.GetCurrentSelection();

            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show(TooltipsLabText.ErrorNoCalloutShapeSelected);
                return;
            }

            foreach (PowerPoint.Shape selectedShape in selection.ShapeRange)
            {
                PowerPoint.Shape triggerShape = CreateTooltip.GenerateTriggerShapeWithReferenceCallout(currentSlide, selectedShape);
                ConvertToTooltip.AddTriggerAnimation(currentSlide, triggerShape, selectedShape);
            }

            if (!this.GetApplication().CommandBars.GetPressedMso(TooltipsLabConstants.AnimationPaneName))
            {
                this.GetApplication().CommandBars.ExecuteMso(TooltipsLabConstants.AnimationPaneName);
            }
        }
    }
}
