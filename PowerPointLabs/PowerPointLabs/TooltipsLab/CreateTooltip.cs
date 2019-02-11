using System;
using System.Windows.Forms;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Views;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.TooltipsLab
{
    static class CreateTooltip
    {

        public static void GenerateCallout(PowerPointSlide currentSlide)
        {
            currentSlide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangularCallout, 0, 0, 100, 100);
        }

        public static void GenerateTriggerShape(PowerPointSlide currentSlide)
        {
            currentSlide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, 100, 100, 10, 10);
        }


    }
}
