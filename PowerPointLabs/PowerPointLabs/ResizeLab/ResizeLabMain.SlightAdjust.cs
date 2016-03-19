using System;
using PowerPointLabs.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    /// <summary>
    /// Implements the Adjust Slightly feature in ResizeLab
    /// </summary>
    public partial class ResizeLabMain
    {
        public const float ResizeFactor = 1;

        #region API

        /// <summary>
        /// Increase the height of shapes
        /// </summary>
        /// <param name="shapes">The shapes to resize</param>
        public void IncreaseHeight(PowerPoint.ShapeRange shapes)
        {
            AdjustHeight(shapes, ResizeFactor);
        }

        /// <summary>
        /// Decrease the height of shapes
        /// </summary>
        /// <param name="shapes">The shapes to resize</param>
        public void DecreaseHeight(PowerPoint.ShapeRange shapes)
        {
            AdjustHeight(shapes, -ResizeFactor);
        }

        /// <summary>
        /// Increase the width of shapes
        /// </summary>
        /// <param name="shapes">The shapes to resize</param>
        public void IncreaseWidth(PowerPoint.ShapeRange shapes)
        {
            AdjustWidth(shapes, ResizeFactor);
        }

        /// <summary>
        /// Decrease the width of shapes
        /// </summary>
        /// <param name="shapes">The shapes to resize</param>
        public void DecreaseWidth(PowerPoint.ShapeRange shapes)
        {
            AdjustWidth(shapes, -ResizeFactor);
        }

        #endregion

        #region Helper functions
        private void AdjustHeight(PowerPoint.ShapeRange shapes, float resizeFactor)
        {
            Action<PPShape> adjustHeight = shape =>
            {
                shape.Top -= resizeFactor;
                shape.AbsoluteHeight += resizeFactor;
            };
            AdjustShape(shapes, adjustHeight);
        }

        private void AdjustWidth(PowerPoint.ShapeRange shapes, float resizeFactor)
        {
            Action<PPShape> adjustWidth = shape => shape.AbsoluteWidth += resizeFactor;
            AdjustShape(shapes, adjustWidth);
        }

        private void AdjustShape(PowerPoint.ShapeRange shapes, Action<PPShape> resizeAction)
        {
            foreach (PowerPoint.Shape shape in shapes)
            {
                var ppShape = new PPShape(shape);
                resizeAction(ppShape);
            }
        }

        #endregion
    }
}
