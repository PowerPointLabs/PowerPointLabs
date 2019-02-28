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

            PowerPoint.Selection selection = this.GetCurrentSelection();

            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("Please select 1 or more shapes as your callout shape.");
                return;
            }

            foreach (PowerPoint.Shape selectedShape in selection.ShapeRange)
            {
                // TODO: Adding a trigger to an existing shape that is already part of a trigger animation as the "callout" (whether created by user or by TooltipsLab)
                //       can possible cause some weird behaviour. For e.g. if user clicked once to show the callout on Trigger A, clicking Trigger B will not hide the callout,
                //       but instead re-show the callout again. Might be a good idea to see if we can detect if the selectedShape is already part of a trigger animation
                //       as a callout, then popup a MessageBox asking if the user is sure he/she wants to continue.

                PowerPoint.Shape triggerShape = CreateTooltip.GenerateTriggerShapeWithReferenceCallout(currentSlide, selectedShape);
                AssignTooltip.AddTriggerAnimation(currentSlide, triggerShape, selectedShape);
            }

            if (!this.GetApplication().CommandBars.GetPressedMso("AnimationCustom"))
            {
                this.GetApplication().CommandBars.ExecuteMso("AnimationCustom");
            }
        }
    }
}