using System.Collections.Generic;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;
using PowerPointLabs.Utils;

namespace PowerPointLabs.PasteLab
{
    static internal class PasteIntoGroup
    {
        public static ShapeRange Execute(PowerPointPresentation presentation, PowerPointSlide slide,
                                    ShapeRange selectedShapes, ShapeRange pastingShapes,
                                    float? posLeft = null, float? posTop = null, int zOrder = 0)
        {
            Shape firstSelectedShape = selectedShapes[1];
            Shape tempShapeForAnimation = slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 0, 1, 1);
            slide.TransferAnimation(firstSelectedShape, tempShapeForAnimation);
            Graphics.MoveZToJustInFront(tempShapeForAnimation, firstSelectedShape);

            string originalGroupName = null;
            if (selectedShapes.Count == 1 && Graphics.IsAGroup(firstSelectedShape))
            {
                originalGroupName = firstSelectedShape.Name;
                selectedShapes = firstSelectedShape.Ungroup();
            }

            // Calculate the center to paste at if not specified
            float selectionLeft = selectedShapes[1].Left;
            float selectionTop = selectedShapes[1].Top;
            float selectionWidth = selectedShapes[1].Width;
            float selectionHeight = selectedShapes[1].Height;
            if (selectedShapes.Count > 1)
            {
                Shape selectionGroup = selectedShapes.Group();
                selectionLeft = selectionGroup.Left;
                selectionTop = selectionGroup.Top;
                selectionWidth = selectionGroup.Width;
                selectionHeight = selectionGroup.Height;
                selectionGroup.Ungroup();
            }
            posLeft = posLeft ?? (selectionLeft + (selectionWidth - pastingShapes[1].Width) / 2);
            posTop = posTop ?? (selectionTop + (selectionHeight - pastingShapes[1].Height) / 2);

            PasteAtCursorPosition.Execute(presentation, slide, pastingShapes, posLeft.Value, posTop.Value);

            List<Shape> shapesToGroupList = new List<Shape>();
            for (int i = 1; i <= selectedShapes.Count; i++)
            {
                shapesToGroupList.Add(selectedShapes[i]);
            }
            for (int i = 1; i <= pastingShapes.Count; i++)
            {
                shapesToGroupList.Add(pastingShapes[i]);
            }

            ShapeRange shapesToGroup = slide.ToShapeRange(shapesToGroupList);
            Shape resultGroup = shapesToGroup.Group();
            resultGroup.Name = originalGroupName ?? resultGroup.Name;
            slide.TransferAnimation(tempShapeForAnimation, resultGroup);
            Graphics.MoveZToJustInFront(resultGroup, tempShapeForAnimation);
            tempShapeForAnimation.Delete();
            if (zOrder == 0)
            {
                pastingShapes.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringToFront);
            }
            else
            {
                //Graphics.MoveZUntilBehind(pastingShapes[1], zOrder);
                //Graphics.MoveZUntilInFront(pastingShapes[1], zOrder);
                new PPShape(pastingShapes[1]).ZOrderPosition = zOrder;
            }

            return slide.ToShapeRange(resultGroup);
        }
    }
}
