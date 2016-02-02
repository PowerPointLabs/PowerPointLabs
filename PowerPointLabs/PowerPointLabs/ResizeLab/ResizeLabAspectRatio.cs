using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    internal static partial class ResizeLabMain
    {
        public static void ChangeShapesAspectRatio(PowerPoint.ShapeRange selectedShapes, bool isAspectRatio)
        {
            try
            {
                if (isAspectRatio)
                {
                    selectedShapes.LockAspectRatio = MsoTriState.msoTrue;
                }
                else
                {
                    selectedShapes.LockAspectRatio = MsoTriState.msoFalse;
                }
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "ChangeShapesAspectRatio");
                throw;
            }
        }

        public static void RestoreAspectRatio(PowerPoint.ShapeRange selectedShapes)
        {
            try
            {
                selectedShapes.ScaleHeight(1, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromTopLeft);
                selectedShapes.ScaleWidth(1, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromTopLeft);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "RestoreAspectRatio");
                throw;
            }
        }
    }
}
