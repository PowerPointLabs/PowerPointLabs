using System.Collections.Generic;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;
using PowerPointLabs.Utils;

namespace PowerPointLabs.MiscFeatures
{
    static internal class MergeIntoGroup
    {
        public static ShapeRange Execute(PowerPointPresentation presentation, PowerPointSlide slide, Selection selection)
        {
            ShapeRange selectedShapes = selection.ShapeRange;
            Shape firstSelectedShape = selectedShapes[1];

            // Temporarily save the animation
            Shape tempShapeForAnimation = slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 0, 1, 1);
            slide.TransferAnimation(firstSelectedShape, tempShapeForAnimation);

            // Merge into one group
            bool isFirstSelectionGroup = false;
            string groupName = firstSelectedShape.Name;

            if (Graphics.IsCorrupted(firstSelectedShape))
            {
                firstSelectedShape = Graphics.CorruptionCorrection(firstSelectedShape, slide);
            }
            if (Graphics.IsAGroup(firstSelectedShape))
            {
                isFirstSelectionGroup = true;

                List<Shape> newShapesList = new List<Shape>();
                ShapeRange ungroupedShapes = firstSelectedShape.Ungroup();
                foreach (Shape shape in ungroupedShapes)
                {
                    newShapesList.Add(shape);
                }
                for (int i = 2; i <= selectedShapes.Count; i++)
                {
                    Shape shape = selectedShapes[i];
                    if (Graphics.IsCorrupted(shape))
                    {
                        shape = Graphics.CorruptionCorrection(shape, slide);
                    }
                    newShapesList.Add(shape);
                }
                selectedShapes = slide.ToShapeRange(newShapesList);
            }

            Shape selectedGroup = selectedShapes.Group();
            selectedGroup.Name = isFirstSelectionGroup ? groupName : selectedGroup.Name;

            // Transfer the animation
            slide.TransferAnimation(tempShapeForAnimation, selectedGroup);
            tempShapeForAnimation.Delete();

            return slide.ToShapeRange(selectedGroup);
        }
    }
}
