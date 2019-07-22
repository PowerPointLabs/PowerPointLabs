using System.Collections.Generic;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ShortcutsLab
{
    internal static class AddIntoGroup
    {
        public static ShapeRange Execute(PowerPointPresentation presentation, PowerPointSlide slide, Selection selection)
        {
            ShapeRange selectedShapes = selection.ShapeRange;
            Shape firstSelectedShape = selectedShapes[1];

            // Temporarily save the animation
            Shape tempShapeForAnimation = slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 0, 1, 1);
            slide.TransferAnimation(firstSelectedShape, tempShapeForAnimation);
            
            // Ungroup first selection and add into list
            string groupName = firstSelectedShape.Name;
            bool isFirstSelectionGroup = false;
            List<Shape> newShapesList = new List<Shape>();

            if (firstSelectedShape.IsCorrupted())
            {
                firstSelectedShape = ShapeUtil.CorruptionCorrection(firstSelectedShape, slide);
            }
            if (firstSelectedShape.IsAGroup())
            {
                isFirstSelectionGroup = true;
                ShapeRange ungroupedShapes = firstSelectedShape.Ungroup();
                foreach (Shape shape in ungroupedShapes)
                {
                    newShapesList.Add(shape);
                }
            }
            else
            {
                newShapesList.Add(firstSelectedShape);
            }
            
            // Add all other selections into list
            for (int i = 2; i <= selectedShapes.Count; i++)
            {
                Shape shape = selectedShapes[i];
                if (shape.IsCorrupted())
                {
                    shape = ShapeUtil.CorruptionCorrection(shape, slide);
                }
                newShapesList.Add(shape);
            }

            // Create the new group from the list
            selectedShapes = slide.ToShapeRange(newShapesList);
            Shape selectedGroup = selectedShapes.Group();
            selectedGroup.Name = isFirstSelectionGroup ? groupName : selectedGroup.Name;

            // Transfer the animation
            slide.TransferAnimation(tempShapeForAnimation, selectedGroup);
            tempShapeForAnimation.Delete();

            return slide.ToShapeRange(selectedGroup);
        }
    }
}
