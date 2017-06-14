using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;

namespace PowerPointLabs.PasteLab
{
    static internal class PasteAtCursorPosition
    {
        public static ShapeRange Execute(PowerPointPresentation presentation, PowerPointSlide slide,
                                        ShapeRange pastingShapes, float positionX, float positionY)
        {
            ShapeRange tempPastingShapes = slide.CopyShapesToSlide(pastingShapes);

            if (tempPastingShapes.Count > 1)
            {
                Shape pastingGroup = tempPastingShapes.Group();
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

            tempPastingShapes.Delete();
            return pastingShapes;
        }
    }
}
