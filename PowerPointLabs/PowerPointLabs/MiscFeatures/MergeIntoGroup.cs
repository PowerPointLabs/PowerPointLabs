using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;

namespace PowerPointLabs.MiscFeatures
{
    static internal class MergeIntoGroup
    {
        public static ShapeRange Execute(PowerPointPresentation presentation, PowerPointSlide slide, Selection selection)
        {
            ShapeRange selectedShapes = selection.ShapeRange;
            Shape tempShapeForAnimation = slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 0, 1, 1);

            slide.TransferAnimation(selectedShapes[1], tempShapeForAnimation);
            Shape selectedGroup = selectedShapes.Group();
            slide.TransferAnimation(tempShapeForAnimation, selectedGroup);

            tempShapeForAnimation.Delete();
            return slide.ToShapeRange(selectedGroup);
        }
    }
}
