﻿using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.CropLab;
using PowerPointLabs.CustomControls;

namespace PowerPointLabs.ActionFramework.CropLab
{
    [ExportActionRibbonId(TextCollection.CropToSlideTag)]
    class CropToSlideActionHandler : CropLabActionHandler
    {

        private static readonly string FeatureName = "Crop To Slide";

        protected override void ExecuteAction(string ribbonId)
        {
            IMessageService cropLabMessageService = MessageServiceFactory.GetCropLabMessageService();
            CropLabErrorHandler errorHandler = CropLabErrorHandler.InitializeErrorHandler(cropLabMessageService);
            if (!IsSelectionShapes(this.GetCurrentSelection()))
            {
                HandleInvalidSelectionError(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, FeatureName, CropLabErrorHandler.SelectionTypeShapeOrPicture, 1, errorHandler);
                return;
            }
            ShapeRange shapeRange = this.GetCurrentSelection().ShapeRange;
            if (shapeRange.Count < 1)
            {
                HandleInvalidSelectionError(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, FeatureName, CropLabErrorHandler.SelectionTypeShapeOrPicture, 1, errorHandler);
                return;
            }
            if (!IsAllPictureOrShape(shapeRange))
            {
                HandleErrorCode(CropLabErrorHandler.ErrorCodeSelectionMustBeShapeOrPicture, FeatureName, errorHandler);
                return;
            }
            float slideWidth = this.GetCurrentPresentation().SlideWidth;
            float slideHeight = this.GetCurrentPresentation().SlideHeight;
            bool hasChange = CropToSlide.Crop(shapeRange, this.GetCurrentSlide(), slideWidth, slideHeight);
            if (!hasChange)
            {
                HandleErrorCode(CropLabErrorHandler.ErrorCodeNoShapeOverBoundary, FeatureName, errorHandler);
            }
        }
        

    }

}
