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
        // To be used for error handling
        internal const int SlightAdjust_MinNoOfShapesRequired = 1;
        internal const string SlightAdjust_FeatureName = "Adjust Slightly";
        internal const string SlightAdjust_ShapeSupport = "object";
        internal static readonly string[] SlightAdjust_ErrorParameters =
        {
            SlightAdjust_FeatureName,
            SlightAdjust_MinNoOfShapesRequired.ToString(),
            SlightAdjust_ShapeSupport
        };

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
                var anchorPoint = GetAnchorPoint(ppShape);

                resizeAction(ppShape);
                AlignShape(ppShape, anchorPoint);
            }
        }

        #endregion
    }
}
