using System;
using PowerPointLabs.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    /// <summary>
    /// SlightAdjust is the partial class of ResizeLabMain.
    /// It handles slight adjustment of the shape's dimensions.
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

        public const float SlightAdjustResizeFactor = 1f;

        #region API

        /// <summary>
        /// Increase the height of shapes
        /// </summary>
        /// <param name="shapes">The shapes to resize</param>
        public void IncreaseHeight(PowerPoint.ShapeRange shapes)
        {
            SlighAdjustHeight(shapes, SlightAdjustResizeFactor);
        }

        /// <summary>
        /// Decrease the height of shapes
        /// </summary>
        /// <param name="shapes">The shapes to resize</param>
        public void DecreaseHeight(PowerPoint.ShapeRange shapes)
        {
            SlighAdjustHeight(shapes, -SlightAdjustResizeFactor);
        }

        /// <summary>
        /// Increase the width of shapes
        /// </summary>
        /// <param name="shapes">The shapes to resize</param>
        public void IncreaseWidth(PowerPoint.ShapeRange shapes)
        {
            SlightAdjustWidth(shapes, SlightAdjustResizeFactor);
        }

        /// <summary>
        /// Decrease the width of shapes
        /// </summary>
        /// <param name="shapes">The shapes to resize</param>
        public void DecreaseWidth(PowerPoint.ShapeRange shapes)
        {
            SlightAdjustWidth(shapes, -SlightAdjustResizeFactor);
        }

        #endregion

        #region Helper functions
        private void SlighAdjustHeight(PowerPoint.ShapeRange shapes, float resizeFactor)
        {
            Action<PPShape> adjustHeight = shape =>
            {
                shape.AbsoluteHeight += resizeFactor;
            };
            SlightAdjustShape(shapes, adjustHeight);
        }

        private void SlightAdjustWidth(PowerPoint.ShapeRange shapes, float resizeFactor)
        {
            Action<PPShape> adjustWidth = shape => shape.AbsoluteWidth += resizeFactor;
            SlightAdjustShape(shapes, adjustWidth);
        }

        private void SlightAdjustShape(PowerPoint.ShapeRange shapes, Action<PPShape> resizeAction)
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
