using System;
using System.Collections.Generic;
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

        public List<float> AdjustProportionallyProportionList;

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

                if (referenceWidth <= 0 || AdjustProportionallyProportionList?.Count != selectedShapes.Count) return;

                for (int i = 1; i < AdjustProportionallyProportionList.Count; i++)
                {
                    var newWidth = referenceWidth * (AdjustProportionallyProportionList[i] / AdjustProportionallyProportionList[0]);
                    var shape = new PPShape(selectedShapes[i + 1]);
                    var anchorPoint = GetAnchorPoint(shape);

                    shape.AbsoluteWidth = newWidth;
                    AlignShape(shape, anchorPoint);
                }
                AdjustProportionallyProportionList = null;
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

                if (referenceHeight <= 0 || AdjustProportionallyProportionList?.Count != selectedShapes.Count) return;

                for (int i = 1; i < AdjustProportionallyProportionList.Count; i++)
                {
                    var newHeight = referenceHeight * (AdjustProportionallyProportionList[i] / AdjustProportionallyProportionList[0]);
                    var shape = new PPShape(selectedShapes[i + 1]);
                    var anchorPoint = GetAnchorPoint(shape);

                    shape.AbsoluteHeight = newHeight;
                    AlignShape(shape, anchorPoint);
                }
                AdjustProportionallyProportionList = null;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AdjustHeightProportionally");
            }
        }
    }
}
