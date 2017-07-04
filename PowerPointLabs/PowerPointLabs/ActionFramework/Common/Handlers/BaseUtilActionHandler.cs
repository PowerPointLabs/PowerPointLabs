﻿using System.Linq;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Interface;

using Office = Microsoft.Office.Core;

namespace PowerPointLabs.ActionFramework.Common.Handlers
{
    abstract class BaseUtilActionHandler : ActionHandler
    {

        protected static bool IsSelectionShapes(Selection selection)
        {
            return selection.Type == PpSelectionType.ppSelectionShapes &&
                    selection.ShapeRange.Count >= 1;
        }

        protected static bool IsAllPictureOrShape(ShapeRange shapeRange)
        {
            return (from Shape shape in shapeRange select shape).All(IsPictureOrShape);
        }

        protected static bool IsAllPicture(ShapeRange shapeRange)
        {
            return (from Shape shape in shapeRange select shape).All(IsPicture);
        }

        protected static bool IsAllPictureWithReferenceObject(ShapeRange shapeRange)
        {
            if (!IsPictureOrShape(shapeRange[1]))
            {
                return false;
            }
            for (int i = 2; i <= shapeRange.Count; i++)
            {
                if (!IsPicture(shapeRange[i]))
                {
                    return false;
                }
            }
            return true;
        }

        protected static bool IsAllShape(ShapeRange shapeRange)
        {
            return (from Shape shape in shapeRange select shape).All(IsShape);
        }

        protected static bool IsPictureOrShape(Shape shape)
        {
            return IsPicture(shape) || IsShape(shape);
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
