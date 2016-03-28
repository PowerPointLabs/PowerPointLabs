using System;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    partial class ResizeLabMain
    {
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

        public float AdjustProportionally_ResizeFactor = 1.5f;

        /// <summary>
        /// Adjust the width of the specified shapes to the resize factor of first
        /// selected shape's width.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void AdjustWidthProportionally(PowerPoint.ShapeRange selectedShapes)
        {
            try
            {
                var referenceWidth = GetReferenceWidth(selectedShapes);

                if (referenceWidth <= 0 || AdjustProportionally_ResizeFactor <= 0) return;

                var newWidth = referenceWidth*AdjustProportionally_ResizeFactor;
                for (int i = 2; i <= selectedShapes.Count; i++)
                {
                    var shape = new PPShape(selectedShapes[i]);
                    var anchorPoint = GetAnchorPoint(shape);

                    shape.AbsoluteWidth = newWidth;
                    AlignShape(shape, anchorPoint);
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AdjustWidthProportionally");
            }
        }

        /// <summary>
        /// Adjust the height of the specified shapes to the resize factor of first
        /// selected shape's height.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void AdjustHeightProportionally(PowerPoint.ShapeRange selectedShapes)
        {
            try
            {
                var referenceHeight = GetReferenceHeight(selectedShapes);

                if (referenceHeight <= 0 || AdjustProportionally_ResizeFactor <= 0) return;

                var newHeight = referenceHeight*AdjustProportionally_ResizeFactor;
                for (int i = 2; i <= selectedShapes.Count; i++)
                {
                    var shape = new PPShape(selectedShapes[i]);
                    var anchorPoint = GetAnchorPoint(shape);

                    shape.AbsoluteHeight = newHeight;
                    AlignShape(shape, anchorPoint);
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AdjustHeightProportionally");
            }
        }
    }
}
