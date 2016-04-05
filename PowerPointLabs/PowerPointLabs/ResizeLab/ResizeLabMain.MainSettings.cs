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

        #region Anchor Point
        public enum AnchorPoint
        {
            TopLeft, TopCenter, TopRight,
            MiddleLeft, Center, MiddleRight,
            BottomLeft, BottomCenter, BottomRight
        }

        public AnchorPoint AnchorPointType { get; set; }

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

            return shape.VisualCenter;
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
                    shape.VisualLeft = anchorPoint.X;
                    shape.VisualTop = anchorPoint.Y;
                    break;
                case AnchorPoint.TopCenter:
                    shape.VisualLeft = anchorPoint.X - shape.AbsoluteWidth / 2;
                    shape.VisualTop = anchorPoint.Y;
                    break;
                case AnchorPoint.TopRight:
                    shape.VisualLeft = anchorPoint.X - shape.AbsoluteWidth;
                    shape.VisualTop = anchorPoint.Y;
                    break;
                case AnchorPoint.MiddleLeft:
                    shape.VisualLeft = anchorPoint.X;
                    shape.VisualTop = anchorPoint.Y - shape.AbsoluteHeight / 2;
                    break;
                case AnchorPoint.Center:
                    shape.VisualLeft = anchorPoint.X - shape.AbsoluteWidth / 2;
                    shape.VisualTop = anchorPoint.Y - shape.AbsoluteHeight / 2;
                    break;
                case AnchorPoint.MiddleRight:
                    shape.VisualLeft = anchorPoint.X - shape.AbsoluteWidth;
                    shape.VisualTop = anchorPoint.Y - shape.AbsoluteHeight / 2;
                    break;
                case AnchorPoint.BottomLeft:
                    shape.VisualLeft = anchorPoint.X;
                    shape.VisualTop = anchorPoint.Y - shape.AbsoluteHeight;
                    break;
                case AnchorPoint.BottomCenter:
                    shape.VisualLeft = anchorPoint.X - shape.AbsoluteWidth / 2;
                    shape.VisualTop = anchorPoint.Y - shape.AbsoluteHeight;
                    break;
                case AnchorPoint.BottomRight:
                    shape.VisualLeft = anchorPoint.X - shape.AbsoluteWidth;
                    shape.VisualTop = anchorPoint.Y - shape.AbsoluteHeight;
                    break;
            }
        }

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
                if (isAspectRatio && selectedShapes.LockAspectRatio == MsoTriState.msoTrue) return;
                if (!isAspectRatio && selectedShapes.LockAspectRatio == MsoTriState.msoFalse) return;

                selectedShapes.LockAspectRatio = isAspectRatio ? MsoTriState.msoTrue : MsoTriState.msoFalse;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "ChangeShapesAspectRatio");
            }
        }

        #endregion

        #region Resize Type
        public enum ResizeBy
        {
            Visual, Actual
        }

        public ResizeBy ResizeType;

        #endregion

    }
}
