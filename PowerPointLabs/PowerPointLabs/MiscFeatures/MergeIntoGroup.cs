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

            Shape tempShapeForAnimation = slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 0, 1, 1);
            slide.TransferAnimation(firstSelectedShape, tempShapeForAnimation);

            Shape selectedGroup = selectedShapes.Group();
            
            slide.TransferAnimation(tempShapeForAnimation, selectedGroup);
            tempShapeForAnimation.Delete();

            return slide.ToShapeRange(selectedGroup);
        }
    }
}
