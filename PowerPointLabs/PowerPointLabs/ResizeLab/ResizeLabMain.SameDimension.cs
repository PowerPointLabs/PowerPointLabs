using System;
using System.Drawing;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    /// <summary>
    /// ResizeLabSameSize is the parital class of ResizeLabMain.
    /// It handles the resizing of the shapes to the same dimension 
    /// (e.g. height, width and both).
    /// </summary>
    public partial class ResizeLabMain
    {
        // To be used for error handling
        internal const int SameDimension_MinNoOfShapesRequired = 2;
        internal const string SameDimension_FeatureName = "Same Dimension";
        internal const string SameDimension_ShapeSupport = "objects";
        internal static readonly string[] SameDimension_ErrorParameters = {
            SameDimension_FeatureName,
            SameDimension_MinNoOfShapesRequired.ToString(),
            SameDimension_ShapeSupport
        };

        public enum SameDimensionAnchor
        {
            TopLeft, TopCenter, TopRight,
            MiddleLeft, Center, MiddleRight,
            BottomLeft, BottomCenter, BottomRight
        }

        public SameDimensionAnchor SameDimensionAnchorType { get; set; }

        /// <summary>
        /// Resize the selected shapes to the same height with the reference to
        /// first selected shape.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void ResizeToSameHeight(PowerPoint.ShapeRange selectedShapes)
        {
            ResizeShapes(selectedShapes, Dimension.Height);
        }

        /// <summary>
        /// Resize the selected shapes to the same width with the reference to
        /// first selected shape.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void ResizeToSameWidth(PowerPoint.ShapeRange selectedShapes)
        {
            ResizeShapes(selectedShapes, Dimension.Width);
        }

        /// <summary>
        /// Resize the selected shapes to the same size (width and height) with 
        /// the reference to first selected shape.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void ResizeToSameHeightAndWidth(PowerPoint.ShapeRange selectedShapes)
        {
            var isAspectRatio = selectedShapes.LockAspectRatio;

            selectedShapes.LockAspectRatio = MsoTriState.msoFalse;
            ResizeShapes(selectedShapes, Dimension.HeightAndWidth);
            selectedShapes.LockAspectRatio = isAspectRatio;
        }

        /// <summary>
        /// Resize the selected shapes according to the set dimension type.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <param name="dimension"></param>
        private void ResizeShapes(PowerPoint.ShapeRange selectedShapes, Dimension dimension)
        {
            try
            {
                var referenceHeight = GetReferenceHeight(selectedShapes);
                var referenceWidth = GetReferenceWidth(selectedShapes);

                if (!IsMoreThanOneShape(selectedShapes, SameDimension_MinNoOfShapesRequired, true, SameDimension_ErrorParameters) 
                    || (referenceHeight < 0) || (referenceWidth < 0))
                {
                    return;
                }

                for (int i = 1; i <= selectedShapes.Count; i++)
                {
                    var shape = new PPShape(selectedShapes[i]);
                    var anchorPoint = GetAnchorPoint(shape);

                    if ((dimension == Dimension.Height) || (dimension == Dimension.HeightAndWidth))
                    {
                        shape.AbsoluteHeight = referenceHeight;
                    }

                    if ((dimension == Dimension.Width) || (dimension == Dimension.HeightAndWidth))
                    {
                        shape.AbsoluteWidth = referenceWidth;
                    }

                    AlignShape(shape, anchorPoint);
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "ResizeShapes");
            }
        }

        /// <summary>
        /// Get the coordinate of anchor point.
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        private PointF GetAnchorPoint(PPShape shape)
        {
            switch (SameDimensionAnchorType)
            {
                case SameDimensionAnchor.TopLeft:
                    return shape.TopLeft;
                case SameDimensionAnchor.TopCenter:
                    return shape.TopCenter;
                case SameDimensionAnchor.TopRight:
                    return shape.TopRight;
                case SameDimensionAnchor.MiddleLeft:
                    return shape.MiddleLeft;
                case SameDimensionAnchor.Center:
                    return shape.Center;
                case SameDimensionAnchor.MiddleRight:
                    return shape.MiddleRight;
                case SameDimensionAnchor.BottomLeft:
                    return shape.BottomLeft;
                case SameDimensionAnchor.BottomCenter:
                    return shape.BottomCenter;
                case SameDimensionAnchor.BottomRight:
                    return shape.BottomRight;
            }

            return shape.Center;
        }

        /// <summary>
        /// Align the shape according to the anchor point given.
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="anchorPoint"></param>
        private void AlignShape(PPShape shape, PointF anchorPoint)
        {
            switch (SameDimensionAnchorType)
            {
                case SameDimensionAnchor.TopLeft:
                    shape.Left = anchorPoint.X;
                    shape.Top = anchorPoint.Y;
                    break;
                case SameDimensionAnchor.TopCenter:
                    shape.Left = anchorPoint.X - shape.AbsoluteWidth/2;
                    shape.Top = anchorPoint.Y;
                    break;
                case SameDimensionAnchor.TopRight:
                    shape.Left = anchorPoint.X - shape.AbsoluteWidth;
                    shape.Top = anchorPoint.Y;
                    break;
                case SameDimensionAnchor.MiddleLeft:
                    shape.Left = anchorPoint.X;
                    shape.Top = anchorPoint.Y - shape.AbsoluteHeight/2;
                    break;
                case SameDimensionAnchor.Center:
                    shape.Left = anchorPoint.X - shape.AbsoluteWidth/2;
                    shape.Top = anchorPoint.Y - shape.AbsoluteHeight/2;
                    break;
                case SameDimensionAnchor.MiddleRight:
                    shape.Left = anchorPoint.X - shape.AbsoluteWidth;
                    shape.Top = anchorPoint.Y - shape.AbsoluteHeight/2;
                    break;
                case SameDimensionAnchor.BottomLeft:
                    shape.Left = anchorPoint.X;
                    shape.Top = anchorPoint.Y - shape.AbsoluteHeight;
                    break;
                case SameDimensionAnchor.BottomCenter:
                    shape.Left = anchorPoint.X - shape.AbsoluteWidth/2;
                    shape.Top = anchorPoint.Y - shape.AbsoluteHeight;
                    break;
                case SameDimensionAnchor.BottomRight:
                    shape.Left = anchorPoint.X - shape.AbsoluteWidth;
                    shape.Top = anchorPoint.Y - shape.AbsoluteHeight;
                    break;
            }
        }

        /// <summary>
        /// Get the height of the reference shape.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <returns></returns>
        private float GetReferenceHeight(PowerPoint.ShapeRange selectedShapes)
        {
            if (selectedShapes.Count > 0)
            {
                return new PPShape(selectedShapes[1]).AbsoluteHeight;
            }
            return -1;
        }

        /// <summary>
        /// Get the width of the reference shape.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <returns></returns>
        private float GetReferenceWidth(PowerPoint.ShapeRange selectedShapes)
        {
            if (selectedShapes.Count > 0)
            {
                return new PPShape(selectedShapes[1]).AbsoluteWidth;
            }
            return -1;
        }
    }
}
