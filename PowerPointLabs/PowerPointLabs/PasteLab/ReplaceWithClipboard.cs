using System.Collections.Generic;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Models;
using PowerPointLabs.SyncLab.ObjectFormats;
using PowerPointLabs.Utils;

namespace PowerPointLabs.PasteLab
{
    static internal class ReplaceWithClipboard
    {
        public static ShapeRange Execute(PowerPointPresentation presentation, PowerPointSlide slide, 
                                        ShapeRange selectedShapes, ShapeRange selectedChildShapes, ShapeRange pastingShapes)
        {
            // ignore height & width, it doesn't always make sense to sync the height & width especially for circles, squares
            List<Format> formatsToIgnore = new List<Format> {new PositionHeightFormat(), new PositionWidthFormat()};
            
            // Replacing individual shape
            if (selectedChildShapes.Count == 0)
            {
                Shape selectedShape = selectedShapes[1];

                Shape pastingShape = pastingShapes[1];
                if (pastingShapes.Count > 1)
                {
                    pastingShape = pastingShapes.Group();
                }
                pastingShape.Left = selectedShape.Left;
                pastingShape.Top = selectedShape.Top;
                ShapeUtil.MoveZToJustInFront(pastingShape, selectedShape);

                slide.DeleteShapeAnimations(pastingShape);
                slide.TransferAnimation(selectedShape, pastingShape);
                ShapeUtil.ApplyAllPossibleFormats(selectedShape, pastingShape, formatsToIgnore);
                // Must remove animations from source shape or else undo will fail
                slide.RemoveAnimationsForShape(selectedShape);
                selectedShape.SafeDelete();

                return slide.ToShapeRange(pastingShape);
            }
            // Replacing shape within a group
            else
            {
                Shape selectedGroup = selectedShapes[1];
                Shape selectedChildShape = selectedChildShapes[1];
                string originalGroupName = selectedGroup.Name;
                int zOrder = selectedChildShape.ZOrderPosition;
                Shape shapeAbove = null;

                float posLeft = selectedChildShape.Left;
                float posTop = selectedChildShape.Top;

                Shape tempShapeForAnimation = slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 0, 1, 1);
                slide.TransferAnimation(selectedGroup, tempShapeForAnimation);

                // Get all siblings of selected child
                List<Shape> selectedGroupShapeList = new List<Shape>();
                for (int i = 1; i <= selectedGroup.GroupItems.Count; i++)
                {
                    Shape shape = selectedGroup.GroupItems.Range(i)[1];
                    if (shape == selectedChildShape)
                    {
                        continue;
                    }
                    selectedGroupShapeList.Add(shape);
                    if (shape.ZOrderPosition - 1 == zOrder)
                    {
                        shapeAbove = shape;
                    }
                }

                // apply all styles from shapes to be pasted, but ignore x,y positions
                // x,y must be applied individually
                // each item replaced has a different positioneach shape in PasteIntoGroup.Execute(...),

                List<Format> positionFormats = new List<Format> {new PositionXFormat(), new PositionYFormat()};
                formatsToIgnore.AddRange(positionFormats);
                
                for (int i = 1; i <= pastingShapes.Count; i++)
                {
                    ShapeUtil.ApplyAllPossibleFormats(selectedChildShape, pastingShapes[i], formatsToIgnore);
                }

                // Remove selected child since it is being replaced
                ShapeRange shapesToGroup = slide.ToShapeRange(selectedGroupShapeList);
                selectedGroup.Ungroup();
                selectedChildShape.SafeDelete();

                ShapeRange result = PasteIntoGroup.Execute(presentation, slide, shapesToGroup, pastingShapes, posLeft, posTop, shapeAbove);
                result[1].Name = originalGroupName;
                slide.TransferAnimation(tempShapeForAnimation, result[1]);

                tempShapeForAnimation.SafeDelete();
                return result;
            }
        }
    }
}
