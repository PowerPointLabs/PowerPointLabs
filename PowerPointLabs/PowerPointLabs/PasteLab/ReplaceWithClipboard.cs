using System;
using System.Collections.Generic;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;
using PowerPointLabs.Utils;

namespace PowerPointLabs.PasteLab
{
    static internal class ReplaceWithClipboard
    {
        public static ShapeRange Execute(PowerPointPresentation presentation, PowerPointSlide slide, Selection selection, ShapeRange pastingShapes)
        {
            // Replacing shape within a group
            if (selection.HasChildShapeRange)
            {
                string uid = DateTime.Now.ToString("ddMMyyyyHHmmssfff");

                Shape selectedGroup = selection.ShapeRange[1];
                Shape selectedChildShape = selection.ChildShapeRange[1];
                selectedChildShape.Tags.Add(PasteLabConstants.ReplaceWithClipboardShapeId, uid);

                float posLeft = selectedChildShape.Left;
                float posTop = selectedChildShape.Top;

                Shape tempShapeForAnimation = slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 0, 1, 1);
                slide.TransferAnimation(selectedGroup, tempShapeForAnimation);

                selectedGroup = Graphics.CorruptionCorrection(selectedGroup, slide);

                List<Shape> selectedGroupShapeList = new List<Shape>();
                int selectedGroupCount = selectedGroup.GroupItems.Count;
                for (int i = 1; i <= selectedGroupCount; i++)
                {
                    Shape shape = selectedGroup.GroupItems.Range(i)[1];
                    if (shape.Tags[PasteLabConstants.ReplaceWithClipboardShapeId].Equals(uid))
                    {
                        continue;
                    }
                    selectedGroupShapeList.Add(shape);
                }

                ShapeRange shapesToGroup = slide.ToShapeRange(selectedGroupShapeList);
                shapesToGroup = slide.CopyShapesToSlide(shapesToGroup);
                selectedGroup.Delete();

                ShapeRange result = PasteIntoGroup.Execute(presentation, slide, shapesToGroup, pastingShapes, posLeft, posTop);
                slide.TransferAnimation(tempShapeForAnimation, result[1]);

                tempShapeForAnimation.Delete();
                return result;
            }
            else // replacing individual shape
            {
                Shape selectedShape = selection.ShapeRange[1];

                Shape pastingShape = pastingShapes[1];
                if (pastingShapes.Count > 1)
                {
                    pastingShape = pastingShapes.Group();
                }
                pastingShape.Left = selectedShape.Left;
                pastingShape.Top = selectedShape.Top;

                slide.DeleteShapeAnimations(pastingShape);
                slide.TransferAnimation(selectedShape, pastingShape);
                selectedShape.Delete();

                return slide.ToShapeRange(pastingShape);
            }
        }
    }
}
