using System;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.CropLab;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("CropToSlideButton")]
    class CropToSlideActionHandler : CropLabActionHandler
    {

        private static readonly string FeatureName = "Crop To Slide";

        protected override void ExecuteAction(string ribbonId)
        {
            CropLabErrorHandler errorHandler = CropLabErrorHandler.InitializeErrorHandler(CropLabUIControl.GetSharedInstance());
            if (!VerifyIsSelectionValid(this.GetCurrentSelection()))
            {
                HandleInvalidSelectionError(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, FeatureName, CropLabErrorHandler.SelectionTypePicture, 1, errorHandler);
                return;
            }
            ShapeRange shapeRange = this.GetCurrentSelection().ShapeRange;
            if (shapeRange.Count < 1)
            {
                HandleInvalidSelectionError(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, FeatureName, CropLabErrorHandler.SelectionTypePicture, 1, errorHandler);
                return;
            }
            if (!IsPictureForSelection(shapeRange))
            {
                HandleErrorCodeIfRequired(CropLabErrorHandler.ErrorCodeSelectionMustBePicture, FeatureName, errorHandler);
                return;
            }
            float slideWidth = this.GetCurrentPresentation().SlideWidth;
            float slideHeight = this.GetCurrentPresentation().SlideHeight;
            CropToSlide.CropSelection(shapeRange, this.GetCurrentSlide(), slideWidth, slideHeight);
        }
        

    }

}
