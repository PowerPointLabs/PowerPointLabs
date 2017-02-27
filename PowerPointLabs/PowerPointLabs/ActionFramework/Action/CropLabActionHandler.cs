using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.CropLab;

using Office = Microsoft.Office.Core;

namespace PowerPointLabs.ActionFramework.Action
{
    abstract class CropLabActionHandler : ActionHandler
    {

        protected bool VerifyIsSelectionValid(Selection selection)
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

        protected static void HandleErrorCodeIfRequired(int errorCode, string featureName, CropLabErrorHandler errorHandler)
        {
            if (errorHandler == null)
            {
                return;
            }
            errorHandler.ProcessErrorCode(errorCode, featureName);
        }

        protected static void HandleInvalidSelectionError(int errorCode, string featureName, string validSelectionType, int validSelectionMinCount, CropLabErrorHandler errorHandler)
        {
            if (errorHandler == null)
            {
                return;
            }
            errorHandler.ProcessErrorCode(errorCode, featureName, validSelectionType, validSelectionMinCount);
        }
    }
}
