using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;
using PowerPointLabs.Utils;

namespace PowerPointLabs.MiscFeatures
{
    static internal class MergeIntoGroup
    {
        public static void Execute(PowerPointPresentation presentation, PowerPointSlide slide, Selection selection)
        {
            ShapeRange selectedShapes = selection.ShapeRange;
            Shape firstSelectedShape = selectedShapes[1];

            string originalGroupName = null;
            if (Graphics.IsAGroup(firstSelectedShape))
            {
                originalGroupName = firstSelectedShape.Name;
            }

            ShapeRange newGroupShapes = slide.CloneShapeFromRange(selectedShapes, firstSelectedShape);
            Shape newGroup = newGroupShapes.Group();
            newGroup.Name = originalGroupName ?? newGroup.Name;
            slide.TransferAnimation(firstSelectedShape, newGroup);

            firstSelectedShape.Delete();
        }
    }
}
