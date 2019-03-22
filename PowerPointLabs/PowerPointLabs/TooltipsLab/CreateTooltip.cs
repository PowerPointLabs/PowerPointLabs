using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.Utility;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.TooltipsLab
{
    internal static class CreateTooltip
    {

        // Generate a trigger shape directly in the center bottom position of the specified callout shape.
        public static PowerPoint.Shape GenerateTriggerShapeWithReferenceCallout(PowerPointSlide currentSlide, PowerPoint.Shape callout)
        {
            float left = ShapeUtil.GetCenterPoint(callout).X - TooltipsLabConstants.TriggerShapeDefaultWidth / 2;
            float top = ShapeUtil.GetBottom(callout) + TooltipsLabConstants.TriggerShapeAndCalloutSpacing;

            PowerPoint.Shape triggerShape = currentSlide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, 
                left, 
                top, 
                TooltipsLabConstants.TriggerShapeDefaultWidth, 
                TooltipsLabConstants.TriggerShapeDefaultHeight);

            return triggerShape;
        }

        public static PowerPoint.Shape GenerateCalloutWithReferenceTriggerShape(PowerPointSlide currentSlide, PowerPoint.Shape triggerShape)
        {
            float midpointX = ShapeUtil.GetMidpointX(triggerShape);


            PowerPoint.Shape callout = currentSlide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangularCallout,
                midpointX - TooltipsLabConstants.CalloutShapeDefaultWidth/2 + (float)(TooltipsLabConstants.CalloutArrowheadHorizontalAdjustment * TooltipsLabConstants.CalloutShapeDefaultWidth),
                triggerShape.Top - (float)(TooltipsLabConstants.CalloutArrowheadVerticalAdjustment * TooltipsLabConstants.CalloutShapeDefaultHeight) - TooltipsLabConstants.TriggerShapeAndCalloutSpacing,
                TooltipsLabConstants.CalloutShapeDefaultWidth,
                TooltipsLabConstants.CalloutShapeDefaultHeight);
            callout.TextEffect.Alignment = MsoTextEffectAlignment.msoTextEffectAlignmentCentered;
            ShapeUtility.FormatCalloutToDefaultStyle(callout);

            return callout;
        }

        public static PowerPoint.Shape GenerateTriggerShape(PowerPointSlide currentSlide)
        {
            PowerPoint.Shape triggerShape = currentSlide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, 
                TooltipsLabConstants.TriggerShapeDefaultLeft, 
                TooltipsLabConstants.TriggerShapeDefaultTop, 
                TooltipsLabConstants.TriggerShapeDefaultWidth, 
                TooltipsLabConstants.TriggerShapeDefaultHeight);
            FormatTriggerShapeToDefaultStyle(triggerShape);
            return triggerShape;
        }

        private static void FormatTriggerShapeToDefaultStyle(PowerPoint.Shape triggerShape)
        {
            triggerShape.TextFrame.TextRange.Font.Size = 16;
            //triggerShape.Fill.ForeColor.RGB = 0;
            triggerShape.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightGray);
            triggerShape.Line.Transparency = 1.0f;
            triggerShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.White);
            triggerShape.TextFrame.TextRange.Text = "?";
        }

    }
}
