using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    /// <summary>
    /// ResizeLabAspectRatio is the parital class of ResizeLabMain.
    /// It controls the related actions of aspect ratio according to
    /// the selection.
    /// </summary>
    internal partial class ResizeLabMain
    {
        public void ChangeShapesAspectRatio(PowerPoint.ShapeRange selectedShapes, bool isAspectRatio)
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

        public void RestoreAspectRatio(PowerPoint.ShapeRange selectedShapes)
        {
            try
            {
                //selectedShapes.ScaleHeight(1, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromTopLeft);
                //selectedShapes.ScaleWidth(1, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromTopLeft);

                var scaleHeight = GetScaleHeight(selectedShapes);
                var scaleWidth = GetScaleWidth(selectedShapes);
                var maximumScale = Math.Max(scaleHeight, scaleWidth);

                selectedShapes.ScaleHeight(maximumScale, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromTopLeft);
                selectedShapes.ScaleWidth(maximumScale, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromTopLeft);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "RestoreAspectRatio");
                throw;
            }
        }

        private float GetScaleHeight(PowerPoint.ShapeRange selectedShapes)
        {
            var currentHeight = selectedShapes.Height;

            selectedShapes.ScaleHeight(1, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromTopLeft);
            var originalHeight = selectedShapes.Height;

            if (Math.Abs(originalHeight - 0) < 0.001)
            {
                return 1;
            }
            return currentHeight/originalHeight;
        }

        private float GetScaleWidth(PowerPoint.ShapeRange selectedShapes)
        {
            var currentWidth = selectedShapes.Width;

            selectedShapes.ScaleWidth(1, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromTopLeft);
            var originalWidth = selectedShapes.Width;

            if (Math.Abs(originalWidth - 0) < 0.001)
            {
                return 1;
            }
            return currentWidth/originalWidth;
        }
    }
}
