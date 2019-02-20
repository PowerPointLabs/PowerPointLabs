using System;
using System.Windows.Forms;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.TooltipsLab
{
    internal static class CreateTooltip
    {

        public static PowerPoint.Shape GenerateCalloutWithReferenceTriggerShape(PowerPointSlide currentSlide, PowerPoint.Shape triggerShape)
        {
            float midpointX = ShapeUtil.GetMidpointX(triggerShape);

            // TODO: To be moved to TooltipsLabConstants in the future.
            int height = 100;
            int width = 150;

            // TODO: 0.20833 and 1.125 and 10 should be moved to TooltipsLabConstants in the future.
            // Explanation for the choice of constants:
            // - 0.20833 is the horizontal percentage adjustment of the arrowhead of the callout.
            //   We position the callout with middle alignment to the trigger shape, then shift it
            //   back to the right by 20.833% of the callout's width to align the arrowhead with the trigger shape.
            // - 1.125 is the vertical percentage adjustment of the arrowhead of the callout.
            //   Same explanation as the horizontal adjustment, just that this is for the height.
            // - 10 is for extra padding between the arrowhead of the callout and the trigger shape.
            PowerPoint.Shape callout = currentSlide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangularCallout,
                midpointX - width/2 + (float)(0.20833 * width),
                triggerShape.Top - (float)(1.125 * height) - 10,
                width,
                height);

            return callout;
        }

        public static PowerPoint.Shape GenerateTriggerShape(PowerPointSlide currentSlide)
        {
            // TODO: These are all hardcoded values for now. They should be replaced in the future.
            PowerPoint.Shape triggerShape = currentSlide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, 200, 200, 25, 25);
            return triggerShape;
        }


    }
}
