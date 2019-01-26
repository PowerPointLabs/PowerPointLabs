using System;
using System.Drawing;

using Microsoft.Office.Core;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    /// <summary>
    /// MainSettings is the partial class of ResizeLabMain.
    /// It controls the settings of related actions in Resize Lab.
    /// </summary>
    public partial class ResizeLabMain
    {

        #region Anchor Point Fields
        public enum AnchorPoint
        {
            TopLeft, TopCenter, TopRight,
            MiddleLeft, Center, MiddleRight,
            BottomLeft, BottomCenter, BottomRight
        }

        public AnchorPoint AnchorPointType { get; set; }
        #endregion

        #region Aspect Ratio

        /// <summary>
        /// Unlocks and locks the aspect ratio of particular period of time.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <param name="isAspectRatio"></param>
        public void ChangeShapesAspectRatio(PowerPoint.ShapeRange selectedShapes, bool isAspectRatio)
        {
            try
            {
                if (isAspectRatio && selectedShapes.LockAspectRatio == MsoTriState.msoTrue)
                {
                    return;
                }

                if (!isAspectRatio && selectedShapes.LockAspectRatio == MsoTriState.msoFalse)
                {
                    return;
                }

                selectedShapes.LockAspectRatio = isAspectRatio ? MsoTriState.msoTrue : MsoTriState.msoFalse;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "ChangeShapesAspectRatio");
            }
        }

        #endregion

        #region Anchor Point Methods
        /// <summary>
        /// Get the coordinate of anchor point.
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        private PointF GetVisualAnchorPoint(PPShape shape)
        {
            switch (AnchorPointType)
            {
                case AnchorPoint.TopLeft:
                    return shape.VisualTopLeft;
                case AnchorPoint.TopCenter:
                    return shape.VisualTopCenter;
                case AnchorPoint.TopRight:
                    return shape.VisualTopRight;
                case AnchorPoint.MiddleLeft:
                    return shape.VisualMiddleLeft;
                case AnchorPoint.Center:
                    return shape.VisualCenter;
                case AnchorPoint.MiddleRight:
                    return shape.VisualMiddleRight;
                case AnchorPoint.BottomLeft:
                    return shape.VisualBottomLeft;
                case AnchorPoint.BottomCenter:
                    return shape.VisualBottomCenter;
                case AnchorPoint.BottomRight:
                    return shape.VisualBottomRight;
            }

            return shape.VisualTopLeft;
        }

        private PointF GetActualAnchorPoint(PPShape shape)
        {
            switch (AnchorPointType)
            {
                case AnchorPoint.TopLeft:
                    return shape.ActualTopLeft;
                case AnchorPoint.TopCenter:
                    return shape.ActualTopCenter;
                case AnchorPoint.TopRight:
                    return shape.ActualTopRight;
                case AnchorPoint.MiddleLeft:
                    return shape.ActualMiddleLeft;
                case AnchorPoint.Center:
                    return shape.ActualCenter;
                case AnchorPoint.MiddleRight:
                    return shape.ActualMiddleRight;
                case AnchorPoint.BottomLeft:
                    return shape.ActualBottomLeft;
                case AnchorPoint.BottomCenter:
                    return shape.ActualBottomCenter;
                case AnchorPoint.BottomRight:
                    return shape.ActualBottomRight;
            }
            return shape.ActualTopLeft;
        }

        /// <summary>
        /// Align the shape according to the anchor point given.
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="anchorPoint"></param>
        private void AlignVisualShape(PPShape shape, PointF anchorPoint)
        {
            switch (AnchorPointType)
            {
                case AnchorPoint.TopLeft:
                    shape.VisualTopLeft = anchorPoint;
                    break;
                case AnchorPoint.TopCenter:
                    shape.VisualTopCenter = anchorPoint;
                    break;
                case AnchorPoint.TopRight:
                    shape.VisualTopRight = anchorPoint;
                    break;
                case AnchorPoint.MiddleLeft:
                    shape.VisualMiddleLeft = anchorPoint;
                    break;
                case AnchorPoint.Center:
                    shape.VisualCenter = anchorPoint;
                    break;
                case AnchorPoint.MiddleRight:
                    shape.VisualMiddleRight = anchorPoint;
                    break;
                case AnchorPoint.BottomLeft:
                    shape.VisualBottomLeft = anchorPoint;
                    break;
                case AnchorPoint.BottomCenter:
                    shape.VisualBottomCenter = anchorPoint;
                    break;
                case AnchorPoint.BottomRight:
                    shape.VisualBottomRight = anchorPoint;
                    break;
            }
        }

        private void AlignActualShape(PPShape shape, PointF anchorPoint)
        {
            switch (AnchorPointType)
            {
                case AnchorPoint.TopLeft:
                    shape.ActualTopLeft = anchorPoint;
                    break;
                case AnchorPoint.TopCenter:
                    shape.ActualTopCenter = anchorPoint;
                    break;
                case AnchorPoint.TopRight:
                    shape.ActualTopRight = anchorPoint;
                    break;
                case AnchorPoint.MiddleLeft:
                    shape.ActualMiddleLeft = anchorPoint;
                    break;
                case AnchorPoint.Center:
                    shape.ActualCenter = anchorPoint;
                    break;
                case AnchorPoint.MiddleRight:
                    shape.ActualMiddleRight = anchorPoint;
                    break;
                case AnchorPoint.BottomLeft:
                    shape.ActualBottomLeft = anchorPoint;
                    break;
                case AnchorPoint.BottomCenter:
                    shape.ActualBottomCenter = anchorPoint;
                    break;
                case AnchorPoint.BottomRight:
                    shape.ActualBottomRight = anchorPoint;
                    break;
            }
        }

        #endregion

        #region Resize Type

        public enum ResizeBy
        {
            Visual,
            Actual
        }

        public ResizeBy ResizeType;

        #endregion
    }
}
