using System.Collections.Generic;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;

namespace PowerPointLabs.PasteLab
{
    static internal class ReplaceWithClipboard
    {
        public static ShapeRange Execute(PowerPointPresentation presentation, PowerPointSlide slide, Selection selection, ShapeRange pastingShapes)
        {
            Shape selectedShape = selection.ShapeRange[1];

            if (selection.HasChildShapeRange)
            {
                selectedShape = selection.ChildShapeRange[1];
                Shape tempSelectedGroup = slide.CopyShapeToSlide(selectedShape.ParentGroup);
                slide.DeleteShapeAnimations(tempSelectedGroup);
                slide.TransferAnimation(selectedShape.ParentGroup, tempSelectedGroup);

                float posLeft = selectedShape.Left;
                float posTop = selectedShape.Top;
                selectedShape.Delete();

                ShapeRange result = PasteIntoGroup.Execute(presentation, slide, selection.ShapeRange, pastingShapes, posLeft, posTop);
                slide.TransferAnimation(tempSelectedGroup, result[1]);

                tempSelectedGroup.Delete();
                return result;
            }

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
