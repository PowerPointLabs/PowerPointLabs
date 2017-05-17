using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.CropLab;
using PowerPointLabs.CustomControls;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("MoveCropShapeButton")]
    class MoveCropShapeButtonActionHandler : CropLabActionHandler
    {
        private static readonly string FeatureName = "Crop To Shape";

        protected override void ExecuteAction(string ribbonId)
        {
            IMessageService cropLabMessageService = MessageServiceFactory.GetCropLabMessageService();
            CropLabErrorHandler errorHandler = CropLabErrorHandler.InitializeErrorHandler(cropLabMessageService);
            if (!IsSelectionShapes(this.GetCurrentSelection()))
            {
                HandleInvalidSelectionError(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, FeatureName, CropLabErrorHandler.SelectionTypeShape, 1, errorHandler);
                return;
            }
            ShapeRange shapeRange = this.GetCurrentSelection().ShapeRange;
            if (shapeRange.Count < 1)
            {
                HandleInvalidSelectionError(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, FeatureName, CropLabErrorHandler.SelectionTypeShape, 1, errorHandler);
                return;
            }
            if (!IsAllShape(shapeRange))
            {
                HandleErrorCode(CropLabErrorHandler.ErrorCodeSelectionMustBeShape, FeatureName, errorHandler);
                return;
            }
            try
            {
                CropToShape.Crop(this.GetCurrentSlide(), this.GetCurrentSelection());
            }
            catch (CropLabException)
            {
                HandleErrorCode(CropLabErrorHandler.ErrorCodeUndefined, FeatureName, errorHandler);
            }
        }
    }
}
