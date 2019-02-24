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
    [ExportActionRibbonId(TooltipsLabText.CreateCalloutTag)]
    class CreateCalloutActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            PowerPointSlide currentSlide = this.GetCurrentSlide();

            PowerPoint.Selection selection = this.GetCurrentSelection();

            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("Please select 1 or more shapes as your trigger shape.");
                return;
            }

            foreach (PowerPoint.Shape selectedShape in selection.ShapeRange)
            {
                PowerPoint.Shape callout = CreateTooltip.GenerateCalloutWithReferenceTriggerShape(currentSlide, selectedShape);
                PowerPoint.Shape calloutGroup = AddTextbox.AddTextboxToCallout(currentSlide, callout);
                AssignTooltip.AddTriggerAnimation(currentSlide, selectedShape, calloutGroup);
            }
            
            if (!this.GetApplication().CommandBars.GetPressedMso("AnimationCustom"))
            {
                this.GetApplication().CommandBars.ExecuteMso("AnimationCustom");
            }
        }
    }
}
