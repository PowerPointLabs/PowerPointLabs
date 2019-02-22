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
                MessageBox.Show("Please select 1 shape as your callout.");
                return;
            }

            foreach (PowerPoint.Shape selectedShape in this.GetAddIn().Application.ActiveWindow.Selection.ShapeRange)
            {
                if (selectedShape.Name.StartsWith("TooltipsLabCallout"))
                {
                    MessageBoxResult result = MessageBox.Show(
                        "The selected shape(s) is/are already associated with a trigger shape(s). Creating a new trigger(s) will invalidate the previous trigger(s). \n\nDo you want to continue?",
                        "Existing trigger detected",
                        MessageBoxButton.YesNo);
                    if (result == MessageBoxResult.No)
                    {
                        return;
                    }
                }

                PowerPoint.Shape triggerShape = CreateTooltip.GenerateTriggerShapeWithReferenceCallout(currentSlide, selectedShape);
                PowerPoint.Shape calloutGroup = AddTextbox.AddTextboxToCallout(currentSlide, selectedShape);
                triggerShape.Name = "TooltipsLabTrigger" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
                calloutGroup.Name = "TooltipsLabCallout" + DateTime.Now.ToString("yyyyMMddHHmmssffff");

                if (selectedShape.Name.StartsWith("TooltipsLabTrigger"))
                {
                    calloutGroup.Name = "TooltipsLabTrigger" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
                   //currentSlide.TimeLine.InteractiveSequences.Get
                } 

                
                AssignTooltip.AddTriggerAnimation(currentSlide, triggerShape, calloutGroup);
            }

            if (!this.GetApplication().CommandBars.GetPressedMso("AnimationCustom"))
            {
                this.GetApplication().CommandBars.ExecuteMso("AnimationCustom");
            }
        }
    }
}
