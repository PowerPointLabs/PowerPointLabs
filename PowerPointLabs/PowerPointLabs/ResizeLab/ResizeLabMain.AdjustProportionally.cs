using System;
using PowerPointLabs.ActionFramework.Common.Log;
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

        public void AdjustWidthProportionally(PowerPoint.ShapeRange selectedShapes, int factor)
        {
            try
            {

            }
            catch (Exception e)
            {
                Logger.LogException(e, "AdjustWidthProportionally");
            }
        }

        public void AdjustHeightProportionally(PowerPoint.ShapeRange selectedShapes, int factor)
        {
            try
            {

            }
            catch (Exception e)
            {
                Logger.LogException(e, "AdjustHeightProportionally");
            }
        }
    }
}
