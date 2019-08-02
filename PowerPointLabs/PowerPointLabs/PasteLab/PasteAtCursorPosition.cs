using System;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;
using PowerPointLabs.Utils;

namespace PowerPointLabs.PasteLab
{
    static internal class PasteAtCursorPosition
    {
        public static ShapeRange Execute(PowerPointPresentation presentation, PowerPointSlide slide,
                                        ShapeRange pastingShapes, float positionX, float positionY)
        {
            if (pastingShapes.Count > 1)
            {
                // Get Left and Top of pasting shapes as a group
                float pastingGroupLeft = int.MaxValue;
                float pastingGroupTop = int.MaxValue;
                foreach (Shape shape in pastingShapes)
                {
                    pastingGroupLeft = Math.Min(shape.Left, pastingGroupLeft);
                    pastingGroupTop = Math.Min(shape.Top, pastingGroupTop);
                }

                foreach (Shape shape in pastingShapes)
                {
                    if (shape.IsAChild())
                    {
                        continue;
                    }
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
