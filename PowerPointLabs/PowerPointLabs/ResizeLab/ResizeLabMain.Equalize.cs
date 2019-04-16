using System;
using System.Drawing;

using Microsoft.Office.Core;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    /// <summary>
    /// Equalize is the partial class of ResizeLabMain.
    /// It handles the resizing of the shapes to the same dimension 
    /// (e.g. height, width and both).
    /// </summary>
    public partial class ResizeLabMain
    {
        // To be used for error handling
        internal const int Equalize_MinNoOfShapesRequired = 2;
        internal const string Equalize_FeatureName = "Equalize";
        internal const string Equalize_ShapeSupport = "objects";
        internal static readonly string[] Equalize_ErrorParameters =
        {
            Equalize_FeatureName,
            Equalize_MinNoOfShapesRequired.ToString(),
            Equalize_ShapeSupport
        };

        /// <summary>
        /// Resize the selected shapes to the same height with the reference to
        /// first selected shape.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void ResizeToSameHeight(PowerPoint.ShapeRange selectedShapes)
        {
            switch (ResizeType)
            {
                case ResizeBy.Visual:
                    ResizeVisualShapes(selectedShapes, Dimension.Height);
                    break;
                case ResizeBy.Actual:
                    ResizeActualShapes(selectedShapes, Dimension.Height);
                    break;
            }
        }

        /// <summary>
        /// Resize the selected shapes to the same width with the reference to
        /// first selected shape.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void ResizeToSameWidth(PowerPoint.ShapeRange selectedShapes)
        {
            switch (ResizeType)
            {
                case ResizeBy.Visual:
                    ResizeVisualShapes(selectedShapes, Dimension.Width);
                    break;
                case ResizeBy.Actual:
                    ResizeActualShapes(selectedShapes, Dimension.Width);
                    break;
            }
        }

        /// <summary>
        /// Resize the selected shapes to the same size (width and height) with 
        /// the reference to first selected shape.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void ResizeToSameHeightAndWidth(PowerPoint.ShapeRange selectedShapes)
        {
            MsoTriState isAspectRatio = selectedShapes.LockAspectRatio;

            selectedShapes.LockAspectRatio = MsoTriState.msoFalse;
            switch (ResizeType)
            {
                case ResizeBy.Visual:
                    ResizeVisualShapes(selectedShapes, Dimension.HeightAndWidth);
                    break;
                case ResizeBy.Actual:
                    ResizeActualShapes(selectedShapes, Dimension.HeightAndWidth);
                    break;
            }
            selectedShapes.LockAspectRatio = isAspectRatio;
        }

        /// <summary>
        /// Resize the selected shapes according to the set dimension type.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <param name="dimension"></param>
        private void ResizeVisualShapes(PowerPoint.ShapeRange selectedShapes, Dimension dimension)
        {
            try
            {
                float referenceHeight = GetReferenceHeight(selectedShapes);
                float referenceWidth = GetReferenceWidth(selectedShapes);

                if (!IsMoreThanOneShape(selectedShapes, Equalize_MinNoOfShapesRequired, true, Equalize_ErrorParameters) 
                    || (referenceHeight < 0) || (referenceWidth < 0))
                {
                    return;
                }

                for (int i = 1; i <= selectedShapes.Count; i++)
                {
                    PPShape shape = new PPShape(selectedShapes[i]);
                    PointF anchorPoint = GetVisualAnchorPoint(shape);

                    if ((dimension == Dimension.Height) || (dimension == Dimension.HeightAndWidth))
                    {
                        shape.AbsoluteHeight = referenceHeight;
                    }

                    if ((dimension == Dimension.Width) || (dimension == Dimension.HeightAndWidth))
                    {
                        shape.AbsoluteWidth = referenceWidth;
                    }

                    AlignVisualShape(shape, anchorPoint);
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "ResizeVisualShapes");
            }
        }

        private void ResizeActualShapes(PowerPoint.ShapeRange selectedShapes, Dimension dimension)
        {
            try
            {
                float referenceHeight = GetReferenceHeight(selectedShapes);
                float referenceWidth = GetReferenceWidth(selectedShapes);

                if (!IsMoreThanOneShape(selectedShapes, Equalize_MinNoOfShapesRequired, true, Equalize_ErrorParameters)
                    || (referenceHeight < 0) || (referenceWidth < 0))
                {
                    return;
                }

                for (int i = 1; i <= selectedShapes.Count; i++)
                {
                    PPShape shape = new PPShape(selectedShapes[i], false);
                    PointF anchorPoint = GetActualAnchorPoint(shape);

                    if ((dimension == Dimension.Height) || (dimension == Dimension.HeightAndWidth))
                    {
                        shape.ShapeHeight = referenceHeight;
                    }

                    if ((dimension == Dimension.Width) || (dimension == Dimension.HeightAndWidth))
                    {
                        shape.ShapeWidth = referenceWidth;
                    }

                    AlignActualShape(shape, anchorPoint);
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "ResizeActualShapes");
            }
        }
    }
}
