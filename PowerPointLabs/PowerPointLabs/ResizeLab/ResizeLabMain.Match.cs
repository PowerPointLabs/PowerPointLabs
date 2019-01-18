using System;

using Microsoft.Office.Core;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    /// <summary>
    /// Match is the partial class of ResizeLabMain.
    /// It handles the resizing of shape's dimension to its respective
    /// width or height
    /// </summary>
    partial class ResizeLabMain
    {
        // To be used for error handling
        internal const int Match_MinNoOfShapesRequired = 1;
        internal const string Match_FeatureName = "Match";
        internal const string Match_ShapeSupport = "object";
        internal static readonly string[] Match_ErrorParameters =
        {
            Match_FeatureName,
            Match_MinNoOfShapesRequired.ToString(),
            Match_ShapeSupport
        };

        /// <summary>
        /// Resize the height of selected shapes to match their width.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void MatchWidth(PowerPoint.ShapeRange selectedShapes)
        {
            switch (ResizeType)
            {
                case ResizeBy.Visual:
                    MatchVisualWidth(selectedShapes);
                    break;
                case ResizeBy.Actual:
                    MatchActualWidth(selectedShapes);
                    break;
            }
        }

        /// <summary>
        /// Resize the width of selected shapes to match their height.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void MatchHeight(PowerPoint.ShapeRange selectedShapes)
        {
            switch (ResizeType)
            {
                case ResizeBy.Visual:
                    MatchVisualHeight(selectedShapes);
                    break;
                case ResizeBy.Actual:
                    MatchActualHeight(selectedShapes);
                    break;
            }
        }

        /// <summary>
        /// Resize the visual height of selected shapes to match their visual width.
        /// </summary>
        /// <param name="selectedShapes"></param>
        private void MatchVisualWidth(PowerPoint.ShapeRange selectedShapes)
        {
            try
            {
                MsoTriState isAspectRatio = selectedShapes.LockAspectRatio;
                selectedShapes.LockAspectRatio = MsoTriState.msoFalse;

                for (int i = 1; i <= selectedShapes.Count; i++)
                {
                    PPShape shape = new PPShape(selectedShapes[i]);
                    System.Drawing.PointF anchorPoint = GetVisualAnchorPoint(shape);

                    shape.AbsoluteHeight = shape.AbsoluteWidth;
                    AlignVisualShape(shape, anchorPoint);
                }

                selectedShapes.LockAspectRatio = isAspectRatio;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "MatchVisualWidth");
            }
        }

        /// <summary>
        /// Resize the visual width of selected shapes to match their visual height.
        /// </summary>
        /// <param name="selectedShapes"></param>
        private void MatchVisualHeight(PowerPoint.ShapeRange selectedShapes)
        {
            try
            {
                MsoTriState isAspectRatio = selectedShapes.LockAspectRatio;
                selectedShapes.LockAspectRatio = MsoTriState.msoFalse;

                for (int i = 1; i <= selectedShapes.Count; i++)
                {
                    PPShape shape = new PPShape(selectedShapes[i]);
                    System.Drawing.PointF anchorPoint = GetVisualAnchorPoint(shape);

                    shape.AbsoluteWidth = shape.AbsoluteHeight;
                    AlignVisualShape(shape, anchorPoint);
                }

                selectedShapes.LockAspectRatio = isAspectRatio;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "MatchVisualHeight");
            }
        }

        /// <summary>
        /// Resize the actual height of selected shapes to match their actual width.
        /// </summary>
        /// <param name="selectedShapes"></param>
        private void MatchActualWidth(PowerPoint.ShapeRange selectedShapes)
        {
            try
            {
                MsoTriState isAspectRatio = selectedShapes.LockAspectRatio;
                selectedShapes.LockAspectRatio = MsoTriState.msoFalse;

                for (int i = 1; i <= selectedShapes.Count; i++)
                {
                    PPShape shape = new PPShape(selectedShapes[i], false);
                    System.Drawing.PointF anchorPoint = GetActualAnchorPoint(shape);

                    shape.ShapeHeight = shape.ShapeWidth;
                    AlignActualShape(shape, anchorPoint);
                }

                selectedShapes.LockAspectRatio = isAspectRatio;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "MatchActualWidth");
            }
        }

        /// <summary>
        /// Resize the actual width of selected shapes to match their actual height.
        /// </summary>
        /// <param name="selectedShapes"></param>
        private void MatchActualHeight(PowerPoint.ShapeRange selectedShapes)
        {
            try
            {
                MsoTriState isAspectRatio = selectedShapes.LockAspectRatio;
                selectedShapes.LockAspectRatio = MsoTriState.msoFalse;

                for (int i = 1; i <= selectedShapes.Count; i++)
                {
                    PPShape shape = new PPShape(selectedShapes[i], false);
                    System.Drawing.PointF anchorPoint = GetActualAnchorPoint(shape);

                    shape.ShapeWidth = shape.ShapeHeight;
                    AlignActualShape(shape, anchorPoint);
                }

                selectedShapes.LockAspectRatio = isAspectRatio;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "MatchActualHeight");
            }
        }
    }
}
