using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;

namespace PowerPointLabs.MiscFeatures
{
    static internal class MergeIntoGroup
    {
        public static void Execute(PowerPointPresentation presentation, PowerPointSlide slide, Selection selection)
        {
            ShapeRange selectedShapes = selection.ShapeRange;
            Shape firstSelectedShape = selectedShapes[1];

            ShapeRange newGroupShapes = slide.CloneShapeFromRange(selectedShapes, firstSelectedShape);
            slide.TransferAnimation(firstSelectedShape, newGroupShapes.Group());

            firstSelectedShape.Delete();
        }
    }
}
