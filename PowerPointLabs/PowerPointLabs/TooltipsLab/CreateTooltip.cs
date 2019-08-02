using System;

using Microsoft.Office.Core;

using PowerPointLabs.Models;
using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.TooltipsLab
{
    internal static class CreateTooltip
    {

        // Generate a trigger shape directly in the center bottom position of the specified callout shape.
        public static PowerPoint.Shape GenerateTriggerShapeWithReferenceCallout(PowerPointSlide currentSlide, PowerPoint.Shape callout)
        {
            float left = callout.GetCenterPoint().X - TooltipsLabConstants.TriggerShapeDefaultWidth / 2;
            float top = callout.GetBottom() + TooltipsLabConstants.TriggerShapeAndCalloutSpacing;

            PowerPoint.Shape triggerShape = currentSlide.Shapes.AddShape(
                TooltipsLabConstants.TriggerShape, 
                left, 
                top, 
                TooltipsLabConstants.TriggerShapeDefaultWidth, 
                TooltipsLabConstants.TriggerShapeDefaultHeight);
            ShapeUtil.FormatTriggerShapeToDefaultStyle(triggerShape);

            return triggerShape;
        }

        public static PowerPoint.Shape GenerateCalloutWithReferenceTriggerShape(PowerPointSlide currentSlide, PowerPoint.Shape triggerShape)
        {
            float midpointX = triggerShape.GetMidpointX();


            PowerPoint.Shape callout = currentSlide.Shapes.AddShape(
                TooltipsLabSettings.ShapeType,
                midpointX - TooltipsLabConstants.CalloutShapeDefaultWidth/2 + (float)(TooltipsLabConstants.CalloutArrowheadHorizontalAdjustment * TooltipsLabConstants.CalloutShapeDefaultWidth),
                triggerShape.Top - (float)(TooltipsLabConstants.CalloutArrowheadVerticalAdjustment * TooltipsLabConstants.CalloutShapeDefaultHeight) - TooltipsLabConstants.TriggerShapeAndCalloutSpacing,
                TooltipsLabConstants.CalloutShapeDefaultWidth,
                TooltipsLabConstants.CalloutShapeDefaultHeight);
            ShapeUtil.FormatCalloutToDefaultStyle(callout);

            callout.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
            callout.TextEffect.Alignment = MsoTextEffectAlignment.msoTextEffectAlignmentCentered;

            return callout;
        }

        public static PowerPoint.Shape GenerateTriggerShape(
            PowerPointPresentation presentation, PowerPointSlide currentSlide)
        {
            PowerPoint.Shape triggerShape = currentSlide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval,
                TooltipsLabConstants.TriggerShapeDefaultLeft,
                TooltipsLabConstants.TriggerShapeDefaultTop,
                TooltipsLabConstants.TriggerShapeDefaultWidth,
                TooltipsLabConstants.TriggerShapeDefaultHeight);
            ShapeUtil.FormatTriggerShapeToDefaultStyle(triggerShape);

            // Look for a shape on the same position of the same size and type on the same slide
            float blurRadius = Math.Min(TooltipsLabConstants.TriggerShapeDefaultWidth,
                TooltipsLabConstants.TriggerShapeDefaultHeight) / 2;
            triggerShape.TryDisplaceShape(presentation, currentSlide, blurRadius);

            return triggerShape;
        }
    }
}
