using System;

using PowerPointLabs.ActionFramework.Common.Log;
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
        public float SlightAdjustResizeFactor = 1f;

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

        #region API

        /// <summary>
        /// Increase the height of shapes
        /// </summary>
        /// <param name="shapes">The shapes to resize</param>
        public void IncreaseHeight(PowerPoint.ShapeRange shapes)
        {
            switch (ResizeType)
            {
                case ResizeBy.Visual:
                    SlightAdjustVisualHeight(shapes, SlightAdjustResizeFactor);
                    break;
                case ResizeBy.Actual:
                    SlightAdjustActualHeight(shapes, SlightAdjustResizeFactor);
                    break;
            }
        }

        /// <summary>
        /// Decrease the height of shapes
        /// </summary>
        /// <param name="shapes">The shapes to resize</param>
        public void DecreaseHeight(PowerPoint.ShapeRange shapes)
        {
            switch (ResizeType)
            {
                case ResizeBy.Visual:
                    SlightAdjustVisualHeight(shapes, -SlightAdjustResizeFactor);
                    break;
                case ResizeBy.Actual:
                    SlightAdjustActualHeight(shapes, -SlightAdjustResizeFactor);
                    break;
            }
        }

        /// <summary>
        /// Increase the width of shapes
        /// </summary>
        /// <param name="shapes">The shapes to resize</param>
        public void IncreaseWidth(PowerPoint.ShapeRange shapes)
        {
            switch (ResizeType)
            {
                case ResizeBy.Visual:
                    SlightAdjustVisualWidth(shapes, SlightAdjustResizeFactor);
                    break;
                case ResizeBy.Actual:
                    SlightAdjustActualWidth(shapes, SlightAdjustResizeFactor);
                    break;
            }
        }

        /// <summary>
        /// Decrease the width of shapes
        /// </summary>
        /// <param name="shapes">The shapes to resize</param>
        public void DecreaseWidth(PowerPoint.ShapeRange shapes)
        {
            switch (ResizeType)
            {
                case ResizeBy.Visual:
                    SlightAdjustVisualWidth(shapes, -SlightAdjustResizeFactor);
                    break;
                case ResizeBy.Actual:
                    SlightAdjustActualWidth(shapes, -SlightAdjustResizeFactor);
                    break;
            }
        }

        #endregion

        #region Helper functions

        private void SlightAdjustVisualHeight(PowerPoint.ShapeRange shapes, float resizeFactor)
        {
            Action<PPShape> adjustHeight = shape => shape.AbsoluteHeight += resizeFactor;
            SlightAdjustVisualShape(shapes, adjustHeight);
        }

        private void SlightAdjustVisualWidth(PowerPoint.ShapeRange shapes, float resizeFactor)
        {
            Action<PPShape> adjustWidth = shape => shape.AbsoluteWidth += resizeFactor;
            SlightAdjustVisualShape(shapes, adjustWidth);
        }

        private void SlightAdjustVisualShape(PowerPoint.ShapeRange shapes, Action<PPShape> resizeAction)
        {
            try
            {
                foreach (PowerPoint.Shape shape in shapes)
                {
                    PPShape ppShape = new PPShape(shape);
                    System.Drawing.PointF anchorPoint = GetVisualAnchorPoint(ppShape);

                    resizeAction(ppShape);
                    AlignVisualShape(ppShape, anchorPoint);
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "SlightAdjustVisualShape");
            }
        }

        private void SlightAdjustActualHeight(PowerPoint.ShapeRange shapes, float resizeFactor)
        {
            Action<PPShape> adjustHeight = shape => shape.ShapeHeight += resizeFactor;
            SlightAdjustActualShape(shapes, adjustHeight);
        }

        private void SlightAdjustActualWidth(PowerPoint.ShapeRange shapes, float resizeFactor)
        {
            Action<PPShape> adjustWidth = shape => shape.ShapeWidth += resizeFactor;
            SlightAdjustActualShape(shapes, adjustWidth);
        }

        private void SlightAdjustActualShape(PowerPoint.ShapeRange shapes, Action<PPShape> resizeAction)
        {
            try
            {
                foreach (PowerPoint.Shape shape in shapes)
                {
                    PPShape ppShape = new PPShape(shape, false);
                    System.Drawing.PointF anchorPoint = GetActualAnchorPoint(ppShape);

                    resizeAction(ppShape);
                    AlignActualShape(ppShape, anchorPoint);
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "SlightAdjustActualShape");
            }
        }

        #endregion
    }
}
