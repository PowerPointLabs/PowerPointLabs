using System.Linq;
using System.Text.RegularExpressions;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Interface;

using Office = Microsoft.Office.Core;

namespace PowerPointLabs.ActionFramework.Util
{
    abstract class BaseUtilActionHandler : ActionHandler
    {

        protected static bool VerifyIsSelectionValid(Selection selection)
        {
            return selection.Type == PpSelectionType.ppSelectionShapes;
        }

        protected static bool IsPictureForSelection(ShapeRange shapeRange)
        {
            return (from Shape shape in shapeRange select shape).All(IsPicture);
        }

        protected static bool IsShapeForSelection(ShapeRange shapeRange)
        {
            return (from Shape shape in shapeRange select shape).All(IsShape);
        }

        protected static bool IsPicture(Shape shape)
        {
            return shape.Type == Office.MsoShapeType.msoPicture ||
                   shape.Type == Office.MsoShapeType.msoLinkedPicture;
        }

        protected static bool IsShape(Shape shape)
        {
            return shape.Type == Office.MsoShapeType.msoAutoShape
                || shape.Type == Office.MsoShapeType.msoFreeform
                || shape.Type == Office.MsoShapeType.msoGroup;
        }


    }
}
