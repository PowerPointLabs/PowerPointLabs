using System;
using System.Collections.Generic;

using Microsoft.Office.Core;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    /// <summary>
    /// AdjustProportionally is the partial class of ResizeLabMain.
    /// It handles the resizing of the shapes to the dimension of given
    /// factor of reference shape.
    /// </summary>
    partial class ResizeLabMain
    {
        public List<float> AdjustProportionallyProportionList;

        // To be used for error handling
        internal const int AdjustProportionally_MinNoOfShapesRequired = 2;
        internal const string AdjustProportionally_FeatureName = "Adjust Proportionally";
        internal const string AdjustProportionally_ShapeSupport = "objects";
        internal static readonly string[] AdjustProportionally_ErrorParameters =
        {
            AdjustProportionally_FeatureName,
            AdjustProportionally_MinNoOfShapesRequired.ToString(),
            AdjustProportionally_ShapeSupport
        };

        /// <summary>
        /// Adjust the width of the specified shapes to the resize factor of first
        /// selected shape's width.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void AdjustWidthProportionally(PowerPoint.ShapeRange selectedShapes)
        {
            switch (ResizeType)
            {
                case ResizeBy.Visual:
                    AdjustVisualWidthProportionally(selectedShapes);
                    break;
                case ResizeBy.Actual:
                    AdjustActualWidthProportionally(selectedShapes);
                    break;
            }
        }

        /// <summary>
        /// Adjust the height of the specified shapes to the resize factor of first
        /// selected shape's height.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void AdjustHeightProportionally(PowerPoint.ShapeRange selectedShapes)
        {
            switch (ResizeType)
            {
                case ResizeBy.Visual:
                    AdjustVisualHeightProportionally(selectedShapes);
                    break;
                case ResizeBy.Actual:
                    AdjustActualHeightProportionally(selectedShapes);
                    break;
            }
        }

        public void AdjustAreaProportionally(PowerPoint.ShapeRange selectedShapes)
        {
            MsoTriState isAspectRatio = selectedShapes.LockAspectRatio;
            bool isLockedRatio = isAspectRatio == MsoTriState.msoTrue;

            selectedShapes.LockAspectRatio = MsoTriState.msoFalse;
            AdjustActualAreaProportionally(selectedShapes, isLockedRatio);
            selectedShapes.LockAspectRatio = isAspectRatio;
        }

        /// <summary>
        /// Adjust the visual width of the specified shapes to the resize factors of first
        /// selected shape's visual width.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void AdjustVisualWidthProportionally(PowerPoint.ShapeRange selectedShapes)
        {
            try
            {
                float referenceWidth = GetReferenceWidth(selectedShapes);

                if (referenceWidth <= 0 || AdjustProportionallyProportionList?.Count != selectedShapes.Count)
                {
                    return;
                }

                for (int i = 1; i < AdjustProportionallyProportionList.Count; i++)
                {
                    float newWidth = referenceWidth*
                                   (AdjustProportionallyProportionList[i] / AdjustProportionallyProportionList[0]);
                    PPShape shape = new PPShape(selectedShapes[i + 1]);
                    System.Drawing.PointF anchorPoint = GetVisualAnchorPoint(shape);

                    shape.AbsoluteWidth = newWidth;
                    AlignVisualShape(shape, anchorPoint);
                }
                AdjustProportionallyProportionList = null;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AdjustVisualWidthProportionally");
            }
        }

        /// <summary>
        /// Adjust the visual height of the specified shapes to the resize factor of first
        /// selected shape's visual height.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void AdjustVisualHeightProportionally(PowerPoint.ShapeRange selectedShapes)
        {
            try
            {
                float referenceHeight = GetReferenceHeight(selectedShapes);

                if (referenceHeight <= 0 || AdjustProportionallyProportionList?.Count != selectedShapes.Count)
                {
                    return;
                }

                for (int i = 1; i < AdjustProportionallyProportionList.Count; i++)
                {
                    float newHeight = referenceHeight*
                                    (AdjustProportionallyProportionList[i] / AdjustProportionallyProportionList[0]);
                    PPShape shape = new PPShape(selectedShapes[i + 1]);
                    System.Drawing.PointF anchorPoint = GetVisualAnchorPoint(shape);

                    shape.AbsoluteHeight = newHeight;
                    AlignVisualShape(shape, anchorPoint);
                }
                AdjustProportionallyProportionList = null;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AdjustVisualHeightProportionally");
            }
        }

        /// <summary>
        /// Adjust the actual width of the specified shapes to the resize factor of first
        /// selected shape's actual width.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void AdjustActualWidthProportionally(PowerPoint.ShapeRange selectedShapes)
        {
            try
            {
                float referenceWidth = GetReferenceWidth(selectedShapes);

                if (referenceWidth <= 0 || AdjustProportionallyProportionList?.Count != selectedShapes.Count)
                {
                    return;
                }

                for (int i = 1; i < AdjustProportionallyProportionList.Count; i++)
                {
                    float newWidth = referenceWidth*
                                   (AdjustProportionallyProportionList[i] / AdjustProportionallyProportionList[0]);
                    PPShape shape = new PPShape(selectedShapes[i + 1], false);
                    System.Drawing.PointF anchorPoint = GetActualAnchorPoint(shape);

                    shape.ShapeWidth = newWidth;
                    AlignActualShape(shape, anchorPoint);
                }
                AdjustProportionallyProportionList = null;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AdjustActualWidthProportionally");
            }
        }

        /// <summary>
        /// Adjust the actual height of the specified shapes to the resize factor of first
        /// selected shape's actual height.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void AdjustActualHeightProportionally(PowerPoint.ShapeRange selectedShapes)
        {
            try
            {
                float referenceHeight = GetReferenceHeight(selectedShapes);

                if (referenceHeight <= 0 || AdjustProportionallyProportionList?.Count != selectedShapes.Count)
                {
                    return;
                }

                for (int i = 1; i < AdjustProportionallyProportionList.Count; i++)
                {
                    float newHeight = referenceHeight*
                                    (AdjustProportionallyProportionList[i] / AdjustProportionallyProportionList[0]);
                    PPShape shape = new PPShape(selectedShapes[i + 1], false);
                    System.Drawing.PointF anchorPoint = GetActualAnchorPoint(shape);

                    shape.ShapeHeight = newHeight;
                    AlignActualShape(shape, anchorPoint);
                }
                AdjustProportionallyProportionList = null;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AdjustActualHeightProportionally");
            }
        }

        /// <summary>
        /// Adjust the actual area of the specified shapes to the resize factor of first
        /// selected shape's actual area.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void AdjustActualAreaProportionally(PowerPoint.ShapeRange selectedShapes, bool isLockedRatio)
        {
            try
            {
                float referenceWidth = selectedShapes[1].Width;
                float referenceHeight = selectedShapes[1].Height;
                double referenceArea = (double)referenceWidth * referenceHeight;
                double referenceRatio = (double)referenceHeight / referenceWidth;

                if (referenceWidth <= 0 || referenceHeight <= 0 || AdjustProportionallyProportionList?.Count != selectedShapes.Count)
                {
                    return;
                }

                for (int i = 1; i < AdjustProportionallyProportionList.Count; i++)
                {
                    PPShape shape = new PPShape(selectedShapes[i + 1], false);
                    System.Drawing.PointF anchorPoint = GetActualAnchorPoint(shape);

                    double newArea = referenceArea *
                                    (AdjustProportionallyProportionList[i] / AdjustProportionallyProportionList[0]);

                    if (isLockedRatio)
                    {
                        referenceWidth = shape.ShapeWidth;
                        referenceHeight = shape.ShapeHeight;
                        referenceRatio = (double)referenceHeight / referenceWidth;
                    }

                    float newWidth = (float)Math.Sqrt(newArea / referenceRatio);
                    float newHeight = (float)(newWidth * referenceRatio);
                    
                    shape.ShapeWidth = newWidth;
                    shape.ShapeHeight = newHeight;
                    AlignActualShape(shape, anchorPoint);
                }
                AdjustProportionallyProportionList = null;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AdjustActualAreaProportionally");
            }
        }
    }
}
