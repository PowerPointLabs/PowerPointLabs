using System;

using Microsoft.Office.Core;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    public class Symmetric
    {
        /// <summary>
        /// Symmetrize new shape at the left of original shape.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void SymmetrizeLeft(PowerPoint.ShapeRange selectedShapes)
        {
            Action<PPShape, PPShape> adjustPosition =
                (originalShape, newShape) =>
                {
                    newShape.VisualLeft = originalShape.VisualLeft - originalShape.AbsoluteWidth;
                    newShape.VisualTop = originalShape.VisualTop;
                }; 

            Symmetrize(selectedShapes, MsoFlipCmd.msoFlipHorizontal, adjustPosition);
        }

        /// <summary>
        /// Symmetrize new shape at the right of original shape.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void SymmetrizeRight(PowerPoint.ShapeRange selectedShapes)
        {
            Action<PPShape, PPShape> adjustPosition =
                (originalShape, newShape) =>
                {
                    newShape.VisualLeft = originalShape.VisualLeft + originalShape.AbsoluteWidth;
                    newShape.VisualTop = originalShape.VisualTop;
                };

            Symmetrize(selectedShapes, MsoFlipCmd.msoFlipHorizontal, adjustPosition);
        }

        /// <summary>
        /// Symmetrize new shape at the top of original shape.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void SymmetrizeTop(PowerPoint.ShapeRange selectedShapes)
        {
            Action<PPShape, PPShape> adjustPosition =
                (originalShape, newShape) =>
                {
                    newShape.VisualLeft = originalShape.VisualLeft;
                    newShape.VisualTop = originalShape.VisualTop - originalShape.AbsoluteHeight;
                };

            Symmetrize(selectedShapes, MsoFlipCmd.msoFlipVertical, adjustPosition);
        }

        /// <summary>
        /// Symmetrize new shape at the bottom of original shape.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void SymmetrizeBottom(PowerPoint.ShapeRange selectedShapes)
        {
            Action<PPShape, PPShape> adjustPosition =
                (originalShape, newShape) =>
                {
                    newShape.VisualLeft = originalShape.VisualLeft;
                    newShape.VisualTop = originalShape.VisualTop + originalShape.AbsoluteHeight;
                };

            Symmetrize(selectedShapes, MsoFlipCmd.msoFlipVertical, adjustPosition);
        }

        /// <summary>
        /// Symmetrize the shapes.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <param name="msoFlipCmd"></param>
        /// <param name="adjustPosition"></param>
        private static void Symmetrize(PowerPoint.ShapeRange selectedShapes, MsoFlipCmd msoFlipCmd, Action<PPShape, PPShape> adjustPosition)
        {
            try
            {
                for (int i = 1; i <= selectedShapes.Count; i++)
                {
                    PPShape originalShape = new PPShape(selectedShapes[i]);
                    PPShape newShape = originalShape.Duplicate();

                    newShape.Flip(msoFlipCmd);
                    newShape.Select(MsoTriState.msoFalse);
                    adjustPosition.Invoke(originalShape, newShape);
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Symmetrize");
            }
        }
    }
}
