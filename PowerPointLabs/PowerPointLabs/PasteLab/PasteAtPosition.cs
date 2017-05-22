using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;

namespace PowerPointLabs.PasteLab
{
    static internal class PasteAtPosition
    {
        public static ShapeRange Execute(PowerPointPresentation presentation, PowerPointSlide slide, 
                                        ShapeRange pastingShapes, float positionX, float positionY)
        {
            if (pastingShapes.Count > 1)
            {
                Shape pastingGroup = slide.CopyShapesToSlide(pastingShapes).Group();
                float pastingGroupLeft = pastingGroup.Left;
                float pastingGroupTop = pastingGroup.Top;
                pastingGroup.Delete();

                foreach (Shape shape in pastingShapes)
                {
                    shape.IncrementLeft(positionX - pastingGroupLeft);
                    shape.IncrementTop(positionY - pastingGroupTop);
                }
            }
            else
            {
                pastingShapes[1].Left = positionX;
                pastingShapes[1].Top = positionY;
            }

            return pastingShapes;
        }
    }
}
