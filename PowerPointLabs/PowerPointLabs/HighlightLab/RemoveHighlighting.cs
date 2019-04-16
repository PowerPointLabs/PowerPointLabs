using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;

namespace PowerPointLabs.HighlightLab
{
    public class RemoveHighlighting
    {
        public static void RemoveHighlight(PowerPointSlide currentSlide)
        {
            currentSlide.DeleteIndicator();
            currentSlide.DeleteShapesWithPrefix("PPTLabsHighlightBackgroundShape");
            currentSlide.DeleteShapesWithPrefix("PPTLabsHighlightTextFragmentsShape");
            foreach (Shape shape in currentSlide.Shapes)
            {
                if (shape.Name.Contains("HighlightTextShape"))
                {
                    currentSlide.DeleteShapeAnimations(shape);
                }
            }
        }
    }
}
