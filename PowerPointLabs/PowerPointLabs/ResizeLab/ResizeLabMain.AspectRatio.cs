using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Log;
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
        private const float FloatDiffTolerance = (float) 0.0001;

        /// <summary>
        /// Unlocks and locks the aspect ratio of particular period of time.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <param name="isAspectRatio"></param>
        public void ChangeShapesAspectRatio(PowerPoint.ShapeRange selectedShapes, bool isAspectRatio)
        {
            try
            {
                selectedShapes.LockAspectRatio = isAspectRatio ? MsoTriState.msoTrue : MsoTriState.msoFalse;
            }
            catch (Exception e) 
            {
                Logger.LogException(e, "ChangeShapesAspectRatio");
            }
        }

        /// <summary>
        /// Restores the shapes to their aspect ratio. The longer side will be used first, and checked 
        /// to ensure that its length is within the length of the slide. If it exceeds the slide, 
        /// the shorter side will be used.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <param name="slideHeight"></param>
        /// <param name="slideWidth"></param>
        public void RestoreAspectRatio(PowerPoint.ShapeRange selectedShapes, float slideHeight, float slideWidth)
        {
            try
            {
                for (int i = 1; i <= selectedShapes.Count; i++)
                {
                    var shape = selectedShapes[i];

                    if (!IsPictureOrOLE(shape)) continue;

                    var scaleHeight = GetScaleHeight(shape);
                    var scaleWidth = GetScaleWidth(shape);
                    var maximumScale = Math.Max(scaleHeight, scaleWidth);
                    var minimumScale = Math.Min(scaleHeight, scaleWidth);

                    if (shape.Height*scaleHeight < slideHeight && shape.Width*scaleWidth < slideWidth)
                    {
                        shape.ScaleHeight(maximumScale, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromMiddle);
                        shape.ScaleWidth(maximumScale, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromMiddle);
                    }
                    else
                    {
                        shape.ScaleHeight(minimumScale, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromMiddle);
                        shape.ScaleWidth(minimumScale, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromMiddle);
                    }
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "RestoreAspectRatio");
            }
        }

        /// <summary>
        /// Get the scale height of the shape at current state.
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        private float GetScaleHeight(PowerPoint.Shape shape)
        {
            var currentHeight = shape.Height;

            shape.ScaleHeight(1, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromMiddle);
            var originalHeight = shape.Height;

            if (IsFloatTheSame(originalHeight, 0))
            {
                return 1;
            }
            return currentHeight/originalHeight;
        }

        /// <summary>
        /// Get the scale width of the shape at current state.
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        private float GetScaleWidth(PowerPoint.Shape shape)
        {
            var currentWidth = shape.Width;

            shape.ScaleWidth(1, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromMiddle);
            var originalWidth = shape.Width;

            if (IsFloatTheSame(originalWidth, 0))
            {
                return 1;
            }
            return currentWidth/originalWidth;
        }

        private bool IsFloatTheSame(float toCompare, float reference)
        {
            return Math.Abs(toCompare - reference) < FloatDiffTolerance;
        }

        private bool IsPictureOrOLE(PowerPoint.Shape shape)
        {
            return shape.Type == MsoShapeType.msoPicture || shape.Type == MsoShapeType.msoLinkedPicture ||
                   shape.Type == MsoShapeType.msoEmbeddedOLEObject || shape.Type == MsoShapeType.msoLinkedOLEObject ||
                   shape.Type == MsoShapeType.msoOLEControlObject;
        }
    }
}
