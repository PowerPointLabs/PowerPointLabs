using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Runtime.InteropServices;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{
    class RemoveHighlighting
    {
#pragma warning disable 0618
        public static void RemoveAllHighlighting()
        {
            var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
            currentSlide.DeleteIndicator();
            currentSlide.DeleteShapesWithPrefix("PPTLabsHighlightBackgroundShape");
            currentSlide.DeleteShapesWithPrefix("PPTLabsHighlightTextFragmentsShape");
            foreach (PowerPoint.Shape sh in currentSlide.Shapes)
            {
                if (sh.Name.Contains("HighlightTextShape"))
                {
                    currentSlide.DeleteShapeAnimations(sh);
                }
            }
        }
    }
}
