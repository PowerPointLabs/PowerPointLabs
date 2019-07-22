using System.Collections.Generic;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;

namespace PowerPointLabs.PasteLab
{
    static internal class PasteIntoGroup
    {
        public static ShapeRange Execute(PowerPointPresentation presentation, PowerPointSlide slide,
                                    ShapeRange selectedShapes, ShapeRange pastingShapes,
                                    float? posLeft = null, float? posTop = null, Shape shapeAbove = null)
        {
            Shape firstSelectedShape = selectedShapes[1];
            Shape tempShapeForAnimation = slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 0, 1, 1);
            slide.TransferAnimation(firstSelectedShape, tempShapeForAnimation);
            ShapeUtil.MoveZToJustInFront(tempShapeForAnimation, firstSelectedShape);

            string originalGroupName = null;
            if (selectedShapes.Count == 1 && ShapeUtil.IsAGroup(firstSelectedShape))
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
            ShapeUtil.MoveZToJustInFront(resultGroup, tempShapeForAnimation);
            tempShapeForAnimation.SafeDelete();
            if (shapeAbove == null)
            {
                pastingShapes.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringToFront);
            }
            else
            {
                ShapeUtil.MoveZToJustBehind(pastingShapes[1], shapeAbove);
            }

            return slide.ToShapeRange(resultGroup);
        }
    }
}
